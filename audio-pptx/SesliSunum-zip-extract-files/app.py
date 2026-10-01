"""Offline DOCX narration -> PowerPoint. No network calls."""
from pathlib import Path
import re
import sys
import wave
import zipfile
import xml.etree.ElementTree as ET
import subprocess
import threading
import queue
import os

# Some Windows Python 3.13 virtual environments miss the base Tcl/Tk path.
# Set paths only inside this process; leave valid explicit settings intact.
if sys.platform == 'win32' and sys.prefix != sys.base_prefix:
    tcl_root = Path(sys.base_prefix) / 'tcl'
    for variable, pattern, required in (
        ('TCL_LIBRARY', 'tcl8.*', 'init.tcl'),
        ('TK_LIBRARY', 'tk8.*', 'tk.tcl'),
    ):
        current = os.environ.get(variable)
        if current and (Path(current) / required).is_file():
            continue
        candidates = sorted(p for p in tcl_root.glob(pattern) if (p / required).is_file())
        if candidates:
            os.environ[variable] = str(candidates[-1])

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
HEADER = re.compile(r'^\s*FOLIE\s*[-–—]\s*(\d+)\s*\.?\s*$', re.I)

def parse_docx(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    result, number = {}, None
    for p in root.findall('.//w:body//w:p', NS):
        text = ''.join(t.text or '' for t in p.findall('.//w:t', NS))
        text = text.strip()
        match = HEADER.fullmatch(text)
        if match:
            number = int(match[1])
            if number < 1 or number in result:
                raise ValueError(f'Geçersiz veya tekrarlanan başlık: {text}')
            result[number] = []
        elif text:
            if number is None:
                raise ValueError('İlk FOLIE başlığından önce metin var. Başlıkları kontrol edin.')
            result[number].append(text)
    if not result:
        raise ValueError('FOLIE - 1. biçiminde başlık bulunamadı.')
    joined = {n: '\n'.join(parts) for n, parts in result.items()}
    empty = [n for n, text in joined.items() if not text.strip()]
    if empty:
        raise ValueError(f'Metni boş slaytlar: {empty}')
    return joined

def validate(numbers, count):
    missing = sorted(set(range(1, count + 1)) - set(numbers))
    extra = sorted(set(numbers) - set(range(1, count + 1)))
    if missing or extra:
        raise ValueError(f'Eşleştirme hatası. Eksik slaytlar: {missing}; fazla numaralar: {extra}')

def configure_audio_timing(slide, shape, duration):
    # Insertion/legacy PlayOnEntry can append a play effect after animations.
    # Remove only effects belonging to this newly inserted audio object.
    main = slide.TimeLine.MainSequence
    sequences = [main]
    interactive = slide.TimeLine.InteractiveSequences
    sequences.extend(interactive.Item(i) for i in range(1, interactive.Count + 1))
    for sequence in sequences:
        for i in range(sequence.Count, 0, -1):
            effect = sequence.Item(i)
            if effect.Shape.Id == shape.Id:
                effect.Delete()
    # MediaPlay=83, AnimationLevelNone=0, WithPrevious=2, Index=1.
    # First in the main sequence, with no dependency on an earlier animation.
    play = main.AddEffect(shape, 83, 0, 2, 1)
    play.Timing.TriggerDelayTime = 0.0
    settings = play.EffectInformation.PlaySettings
    settings.PauseAnimation = 0
    settings.HideWhileNotPlaying = -1
    settings.StopAfterSlides = 1
    settings.LoopUntilStopped = 0
    transition = slide.SlideShowTransition
    transition.AdvanceOnTime = -1
    transition.AdvanceOnClick = 0
    # Include the existing transition duration and a small playback margin.
    transition.AdvanceTime = duration + max(0.0, float(transition.Duration)) + 2.0

def narrate(ppt, doc, model, target, log):
    import pythoncom
    import win32com.client
    model = Path(model).resolve()
    if not model.is_file() or not Path(str(model) + '.json').is_file():
        raise ValueError('Yerel .onnx ve aynı adlı .onnx.json model dosyaları gerekli.')
    texts = parse_docx(doc)
    pythoncom.CoInitialize()
    presentation = None
    try:
        # Do not quit PowerPoint: the user may have other presentations open.
        office = win32com.client.Dispatch('PowerPoint.Application')
        presentation = office.Presentations.Open(str(Path(ppt).resolve()), -1, 0, 0)
        validate(texts, presentation.Slides.Count)
        run = Path(target).resolve() / ('SesliSunum_' + __import__('datetime').datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
        run.mkdir(parents=True, exist_ok=False)
        audio = run / 'sesler'
        audio.mkdir()
        for n in range(1, presentation.Slides.Count + 1):
            log(f'Slayt {n}/{presentation.Slides.Count}: ses üretiliyor…')
            wav = audio / f'folie_{n:03}.wav'
            # Text uses stdin, not shell arguments. Explicit local model; no downloader.
            # Parent encodes stdin as UTF-8; Piper must decode it as UTF-8 too.
            # Windows child processes otherwise may use a legacy code page.
            child_env = os.environ.copy()
            child_env['PYTHONIOENCODING'] = 'utf-8'
            child_env['PYTHONUTF8'] = '1'
            proc = subprocess.run([sys.executable, '-X', 'utf8', '-m', 'piper', '-m', str(model), '-f', str(wav)],
                                  input=texts[n], text=True, encoding='utf-8', capture_output=True,
                                  env=child_env, creationflags=subprocess.CREATE_NO_WINDOW, timeout=600)
            if proc.returncode:
                raise RuntimeError(f'Piper ses oluşturamadı (slayt {n}). Kurulum ve modeli kontrol edin.')
            with wave.open(str(wav), 'rb') as sound:
                duration = sound.getnframes() / sound.getframerate()
            slide = presentation.Slides(n)
            shape = slide.Shapes.AddMediaObject2(str(wav), 0, -1, -100, -100, 16, 16)
            shape.Name = f'OfflineNarration_{n}'
            configure_audio_timing(slide, shape, duration)
        presentation.SlideShowSettings.AdvanceMode = 2
        output = run / (Path(ppt).stem + '_sesli.pptx')
        presentation.SaveAs(str(output), 24)
        # Verify every generated sound is embedded in the resulting package.
        with zipfile.ZipFile(output) as z:
            embedded = [i for i in z.infolist() if i.filename.startswith('ppt/media/') and i.filename.endswith('.wav')]
            if len(embedded) < len(texts):
                raise RuntimeError('Ses gömme kontrolü başarısız. Çıktıyı kullanmadan kontrol edin.')
        log(f'Tamamlandı: {output}')
    finally:
        if presentation is not None:
            presentation.Close()
        pythoncom.CoUninitialize()

def main():
    root = tk.Tk()
    root.title('Sesli Sunum — Yerel Almanca Seslendirme')
    root.geometry('800x650')
    panel = ttk.Frame(root, padding=15)
    panel.pack(fill='both', expand=True)
    fields = {}
    def choose(key):
        if key == 'Çıktı klasörü':
            value = filedialog.askdirectory()
        else:
            types = {'PowerPoint': [('PowerPoint', '*.pptx *.ppt')], 'Word': [('Word', '*.docx')],
                     'Almanca ses modeli': [('Piper model', '*.onnx')]}
            value = filedialog.askopenfilename(filetypes=types[key])
        if value:
            fields[key].set(value)
    for key in ('PowerPoint', 'Word', 'Almanca ses modeli', 'Çıktı klasörü'):
        ttk.Label(panel, text=key).pack(anchor='w')
        row = ttk.Frame(panel)
        row.pack(fill='x', pady=(0, 8))
        fields[key] = tk.StringVar()
        ttk.Entry(row, textvariable=fields[key]).pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='Seç', command=lambda k=key: choose(k)).pack(side='right')
    ttk.Label(panel, text='Word başlıkları: FOLIE - 1. / FOLIE - 2.\nHer slayt için metin gerekir. Kaynak dosya değiştirilmez.').pack(anchor='w', pady=8)
    box = tk.Text(panel, wrap='word', height=16)
    box.pack(fill='both', expand=True)
    events = queue.Queue()
    def preview():
        try:
            texts = parse_docx(fields['Word'].get())
            box.delete('1.0', 'end')
            for n, text in texts.items():
                box.insert('end', f'FOLIE {n}\n{text}\n\n')
        except Exception as e:
            messagebox.showerror('Kontrol', str(e))
    def start():
        values = [fields[k].get() for k in fields]
        if not all(values):
            messagebox.showerror('Eksik seçim', 'Dört alanı da doldurun.')
            return
        create.config(state='disabled')
        def worker():
            try:
                narrate(*values, events.put)
            except Exception as e:
                events.put(f'HATA: {e}')
            finally:
                events.put(None)
        threading.Thread(target=worker, daemon=True).start()
    def pump():
        while not events.empty():
            item = events.get()
            if item is None:
                create.config(state='normal')
            else:
                box.insert('end', item + '\n')
                box.see('end')
        root.after(150, pump)
    row = ttk.Frame(panel)
    row.pack(fill='x', pady=10)
    ttk.Button(row, text='Metinleri önizle', command=preview).pack(side='left')
    create = ttk.Button(row, text='Sesli PPTX oluştur', command=start)
    create.pack(side='right')
    def close():
        if str(create['state']) == 'disabled':
            messagebox.showinfo('İşlem sürüyor', 'Dosyanın tamamlanmasını bekleyin.')
        else:
            root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    pump()
    root.mainloop()

if __name__ == '__main__':
    main()
