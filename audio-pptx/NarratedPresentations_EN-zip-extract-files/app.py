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
                raise ValueError(f'Invalid or duplicate heading: {text}')
            result[number] = []
        elif text:
            if number is None:
                raise ValueError('Text appears before the first FOLIE heading. Check the headings.')
            result[number].append(text)
    if not result:
        raise ValueError('No headings in the FOLIE - 1. format were found.')
    joined = {n: '\n'.join(parts) for n, parts in result.items()}
    empty = [n for n, text in joined.items() if not text.strip()]
    if empty:
        raise ValueError(f'Slides with empty narration: {empty}')
    return joined

def validate(numbers, count):
    missing = sorted(set(range(1, count + 1)) - set(numbers))
    extra = sorted(set(numbers) - set(range(1, count + 1)))
    if missing or extra:
        raise ValueError(f'Slide mapping error. Missing slides: {missing}; extra slide numbers: {extra}')

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
        raise ValueError('A local .onnx model and its matching .onnx.json file are required.')
    texts = parse_docx(doc)
    pythoncom.CoInitialize()
    presentation = None
    try:
        # Do not quit PowerPoint: the user may have other presentations open.
        office = win32com.client.Dispatch('PowerPoint.Application')
        presentation = office.Presentations.Open(str(Path(ppt).resolve()), -1, 0, 0)
        validate(texts, presentation.Slides.Count)
        run = Path(target).resolve() / ('NarratedPresentation_' + __import__('datetime').datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
        run.mkdir(parents=True, exist_ok=False)
        audio = run / 'audio'
        audio.mkdir()
        for n in range(1, presentation.Slides.Count + 1):
            log(f'Slide {n}/{presentation.Slides.Count}: generating audio…')
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
                raise RuntimeError(f'Piper could not generate audio for slide {n}. Check the installation and model.')
            with wave.open(str(wav), 'rb') as sound:
                duration = sound.getnframes() / sound.getframerate()
            slide = presentation.Slides(n)
            shape = slide.Shapes.AddMediaObject2(str(wav), 0, -1, -100, -100, 16, 16)
            shape.Name = f'OfflineNarration_{n}'
            configure_audio_timing(slide, shape, duration)
        presentation.SlideShowSettings.AdvanceMode = 2
        output = run / (Path(ppt).stem + '_narrated.pptx')
        presentation.SaveAs(str(output), 24)
        # Verify every generated sound is embedded in the resulting package.
        with zipfile.ZipFile(output) as z:
            embedded = [i for i in z.infolist() if i.filename.startswith('ppt/media/') and i.filename.endswith('.wav')]
            if len(embedded) < len(texts):
                raise RuntimeError('Audio embedding verification failed. Check the output before using it.')
        log(f'Completed: {output}')
    finally:
        if presentation is not None:
            presentation.Close()
        pythoncom.CoUninitialize()

def main():
    from webapp import main as browser_main
    browser_main()

if __name__ == '__main__':
    main()

