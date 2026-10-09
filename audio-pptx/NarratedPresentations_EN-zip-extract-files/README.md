# Narrated Presentations

Create German narration from a Word script and embed it in PowerPoint slides.
Requires Windows, desktop PowerPoint, and 64-bit Python. Python 3.13 is supported.
No Tcl/Tk installation is needed. Try a short sample before processing company files.

## One-time setup

Open PowerShell in this folder and run:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
New-Item -ItemType Directory -Force models
.\.venv\Scripts\python.exe -m piper.download_voices de_DE-thorsten-medium --data-dir models
```

Setup and model download require internet access. No presentation content is processed during setup.
Skip the model download if the model is already in `models`.
For an entirely offline installation, ask IT to prepare the required Python packages and model files
on an approved computer with the matching Windows/Python version.
Piper uses GPL-3.0 licensing. Ask IT to review the software, dependencies, and voice model terms.
The Thorsten model card identifies its training dataset as CC0.

- Software: https://github.com/OHF-Voice/piper1-gpl
- Voice model card: https://huggingface.co/rhasspy/piper-voices/blob/main/de/de_DE/thorsten/medium/MODEL_CARD

## Use

1. Double-click `Start.vbs`. The browser opens without a console window.
2. Select the original PowerPoint file and Word `.docx` script.
3. Select the `.onnx` model in `models`. Its matching `.onnx.json` must be alongside it.
4. Choose an existing local output folder.
5. Click **Preview script** and check the slide mapping.
6. Click **Create narrated PPTX** and watch the progress messages.
7. Open the generated `_narrated.pptx` in PowerPoint and press F5 to test playback.
8. When finished, click **Close application**. Closing the browser tab alone does not stop the server.

If a file picker does not appear, paste the full path into the field without surrounding quotes.
Windows dialog buttons use your Windows display language.
Do not open `index.html` directly: the local Python backend must be running.
If VBS is blocked by company policy, use `Start.cmd` to launch the app and view diagnostic messages.

## Word script format

Each heading must be a separate paragraph. German narration remains in German:

```text
FOLIE - 1.
Guten Tag. Heute stellen wir unser Beispielprojekt vor.

FOLIE - 2.
Im zweiten Schritt betrachten wir die wichtigsten Ergebnisse.

FOLIE - 3.
Vielen Dank für Ihre Aufmerksamkeit.
```

Heading case, spaces around the dash, and the final period are flexible.
Slide numbers refer to physical slide order, including hidden slides.
Missing, extra, duplicate, or empty sections are rejected. No text should precede the first heading.
Body paragraphs and paragraphs in tables are read. Text boxes, headers, and footnotes are excluded.
Convert old `.doc` files to `.docx` in Word first.

## Output and playback

Each run creates a new `NarratedPresentation_...` folder containing the PPTX and an `audio` folder.
Audio is embedded in the PPTX. The source presentation is opened read-only.
The audio play effect is placed first in the animation sequence with zero trigger delay.
Slide timing is audio duration plus the existing transition duration plus a two-second margin.
Advancement uses a timer, rather than an audio-ended event. Verify playback on the target computer.
Existing music or narration is not removed and may overlap. Always use the original unnarrated file.
Click-triggered animations are not synchronized to spoken words.
For MP4 output, use PowerPoint **File > Export > Create a Video**, using recorded timings.
After an error, do not use an incomplete output; resolve the error and run again.

## Local processing and transfer

The browser connects only to a Python server bound to `127.0.0.1` on this computer.
There are no external APIs or model downloads during narration. Office and other software may
have their own network behavior. Test with internet disconnected and use folders outside cloud sync.
Diagnostic messages can include local file paths. System-generated errors may follow the OS language.

To move to another computer, copy this application folder including `models`.
Do not copy `.venv` or `__pycache__`; create a new virtual environment on the target computer.
The parent `work` folder is not needed. This is a Python source package, not a standalone EXE.

## Updating an existing installation

Close the application, then copy the new package files into your existing application folder.
Keep `.venv` and `models`. Use `Start.vbs` from now on.
Old `Baslat.vbs`, `Baslat.cmd`, and `KULLANIM.md` files may be removed if they remain from an older package.
