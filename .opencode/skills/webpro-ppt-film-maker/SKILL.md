---
name: webpro-ppt-film-maker
description: Use ONLY when working on webpro_ppt_film_maker, PowerPoint-to-video workflows, app.py, gui.py, README.md, WORKFLOW.md, MiniMax TTS/Hailuo video generation, OpenAI image generation, ffmpeg composition, or project-specific CLI/GUI debugging.
---

# Webpro PPT Film Maker

Use this skill for this repository's PowerPoint-to-video tool. The project provides a Python CLI and a tkinter GUI for generating template images, scene images, speech audio, video clips, subtitles, composed clips, and final concatenated videos.

## Project Map

- `app.py`: Main CLI implementation and all API/media task logic.
- `gui.py`: tkinter GUI wrapper that builds `app.py` subprocess commands.
- `README.md`: User-facing setup and command examples.
- `WORKFLOW.md`: High-level PowerPoint-to-video production workflow.
- `requirements.txt`: Notes that no third-party Python packages are currently required.

## Environment

Required runtime:

- Python 3.10+
- `ffmpeg` and `ffprobe` available on `PATH`
- `OPENAI_API_KEY_FOR_IMAGE` for OpenAI image generation
- `MINIMAX_API_KEY` for MiniMax TTS and Hailuo video generation

Environment variables:

```sh
export OPENAI_API_KEY_FOR_IMAGE="your OpenAI API key"
export MINIMAX_API_KEY="your MiniMax API key"
```

`OPENAI_API_KEY` is supported only as a legacy fallback for image tasks. Prefer `OPENAI_API_KEY_FOR_IMAGE`.

## Run Commands

Show CLI help:

```sh
python app.py --help
```

Start the GUI:

```sh
python gui.py
```

If GUI startup fails with `_tkinter` or `tkinter` import errors on macOS Homebrew Python, install the matching Tk package and recreate the virtual environment.

```sh
brew install python-tk@3.13
rm -rf .venv
/opt/homebrew/bin/python3.13 -m venv .venv
source .venv/bin/activate
python gui.py
```

## CLI Tasks

The CLI uses `--task` with one of these values:

- `template`: Generate a reusable template image with OpenAI Images API.
- `scene`: Generate a scene image from a PPT slide image and template image.
- `tts`: Generate speech audio with MiniMax TTS.
- `video`: Generate a video clip with MiniMax Hailuo.
- `subtitle`: Burn subtitles into a video with ffmpeg.
- `compose`: Merge video and audio with ffmpeg.
- `concat`: Concatenate clips into the final video.

Template image:

```sh
python app.py --task template \
  --prompt "高質感新聞攝影棚，大型簡報牆，右側主持人區域" \
  --reference ref1.png \
  --output template.png
```

Scene image:

```sh
python app.py --task scene \
  --ppt-image slide1.png \
  --template-image template.png \
  --scene-number 1 \
  --prompt "將 PPT 放入大型簡報牆，保留文字與版面清晰可讀" \
  --output sence1.png
```

Speech audio:

```sh
python app.py --task tts \
  --text-file script.txt \
  --voice English_Graceful_Lady \
  --speed 1.0 \
  --audio-format mp3 \
  --output speaking-sound/sence1.mp3
```

Video clip:

```sh
python app.py --task video \
  --image sence1.png \
  --prompt "鏡頭緩慢推進，畫面自然流動" \
  --duration 6 \
  --resolution 1080P \
  --output videos/sence1.mp4
```

Subtitle burn-in:

```sh
python app.py --task subtitle \
  --video-input videos/sence1.mp4 \
  --subtitle-input subtitles/sence1.srt \
  --output videos/sence1_subtitled.mp4
```

Compose video and audio:

```sh
python app.py --task compose \
  --video-input videos/sence1.mp4 \
  --audio-input speaking-sound/sence1.mp3 \
  --output merged/sence1.mp4
```

Concatenate clips:

```sh
python app.py --task concat --input-dir merged --output final.mp4
```

## Production Workflow

Follow the intended workflow from `WORKFLOW.md`:

1. Create the PowerPoint content and export each slide as PNG or JPEG.
2. Generate and refine a template image with `template`.
3. Generate per-slide scene images with `scene`, naming them `sence1`, `sence2`, `sence3`, etc.
4. Generate narration audio with `tts`.
5. Estimate each video duration from narration audio when needed.
6. Generate scene videos with `video`, optionally using start and end images.
7. Merge each video clip with its narration using `compose`.
8. Concatenate all composed clips using `concat`.

Preserve the existing `sence*` spelling because the current CLI and documentation use it for sorting and examples.

## Implementation Notes

- Keep API logic in `app.py` unless the task is purely GUI behavior.
- Keep `gui.py` as a subprocess wrapper around `app.py`; do not duplicate API or ffmpeg logic in the GUI.
- The GUI should surface stdout and stderr from CLI execution.
- Image tasks use OpenAI and default to `gpt-image-2`.
- Video tasks use MiniMax and default to `MiniMax-Hailuo-02`.
- TTS defaults to `speech-2.8-hd`, `English_Graceful_Lady`, and `mp3`.
- `video --audio-input` can be used to estimate duration via `ffprobe` when `--duration` is not supplied.
- `template` supports 0 to 4 `--reference` images.
- `scene` requires `--ppt-image`; it can also use template and reference images.
- All tasks support `--output-json` for structured output.

## Change Rules

- Read `README.md`, `WORKFLOW.md`, and the relevant section of `app.py` or `gui.py` before changing behavior.
- Prefer small, local edits that preserve the CLI contract documented in `README.md`.
- When changing CLI arguments, update `README.md` and GUI command construction if affected.
- When changing task validation in `app.py`, ensure GUI inputs still map to valid CLI arguments.
- Do not remove legacy `OPENAI_API_KEY` fallback unless explicitly requested.
- Do not rename `sence*` outputs unless doing a coordinated migration across docs, sorting logic, examples, and GUI labels.
- Avoid adding third-party dependencies unless there is a concrete need.

## Verification

Use lightweight checks when API keys or media inputs are unavailable:

```sh
python app.py --help
python -m py_compile app.py gui.py
```

For media tasks, also verify external tools:

```sh
ffmpeg -version
ffprobe -version
```

For GUI work, at minimum verify import/startup behavior where tkinter is available:

```sh
python gui.py
```

Do not run paid API generation tasks unless the user explicitly asks or provides suitable test inputs.
