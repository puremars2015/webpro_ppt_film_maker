# PowerPoint 轉影片工具

這個專案提供兩個入口：

- `app.py`：Python CLI 工具
- `gui.py`：tkinter GUI，透過 subprocess 呼叫 CLI

## 環境需求

- Python 3.10+
- ffmpeg 與 ffprobe
- OpenAI API Key，用於模板與場景圖片生成
- MiniMax API Key，用於 TTS 與 Hailuo 影片生成

若要使用 GUI，Python 必須內建 tkinter 支援。在 macOS 上，若使用 Homebrew Python，可能需要另外安裝對應版本的 Tk 套件。

API Key 預設讀取系統環境變數：

```sh
export OPENAI_API_KEY_FOR_IMAGE="你的 OpenAI API Key"
export MINIMAX_API_KEY="你的 MiniMax API Key"
```

圖片相關任務預設優先讀取 `OPENAI_API_KEY_FOR_IMAGE`；若你仍有舊設定，CLI 也相容讀取 `OPENAI_API_KEY`。

## 執行

查看 CLI 說明：

```sh
python app.py --help
```

啟動 GUI：

```sh
python gui.py
```

### macOS tkinter 問題

如果啟動 GUI 時看到 `No module named '_tkinter'`，代表目前的 Python 不含 Tk 支援。以 Homebrew Python 3.13 為例，可用以下方式處理：

```sh
brew install python-tk@3.13
rm -rf .venv
/opt/homebrew/bin/python3.13 -m venv .venv
source .venv/bin/activate
python gui.py
```

若你改用其他 Python 版本建立虛擬環境，也請確認該版本本身可以成功 `import tkinter`。

## CLI 任務

### 1. 製作模板畫面

呼叫 OpenAI Images API。可不提供參考圖，也可提供 0 到 4 張參考圖。

```sh
python app.py --task template \
  --prompt "高質感新聞攝影棚，大型簡報牆，右側主持人區域" \
  --reference ref2.png \
  --reference ref1.png \
  --output template.png
```

可重複使用 `--reference` 來附加多張參考圖。

### 2. 產生場景圖片

用 PPT 匯出圖片、模板圖片與 prompt 產生 `sence*.png`。

```sh
python app.py --task scene \
  --ppt-image slide1.png \
  --template-image template.png \
  --scene-number 1 \
  --prompt "將 PPT 放入大型簡報牆，保留文字與版面清晰可讀" \
  --output sence1.png
```

### 3. 產生語音檔

呼叫 MiniMax TTS API。

```sh
python app.py --task tts \
  --text "這是一段旁白文稿。" \
  --voice English_Graceful_Lady \
  --speed 1.0 \
  --audio-format mp3 \
  --output speaking-sound/sence1.mp3
```

也可以從文字檔讀取：

```sh
python app.py --task tts --text-file script.txt --output speech.mp3
```

### 4. 產生影片檔

呼叫 MiniMax Hailuo 影片 API。支援起始圖片與可選結束圖片。

```sh
python app.py --task video \
  --image sence1.png \
  --prompt "鏡頭緩慢推進，畫面自然流動" \
  --duration 6 \
  --resolution 1080P \
  --output videos/sence1.mp4
```

若提供 `--audio-input` 且未提供 `--duration`，CLI 會用 ffprobe 依音檔長度估算秒數。

### 5. 燒入字幕

使用 ffmpeg 將字幕檔燒入影片。

```sh
python app.py --task subtitle \
  --video-input videos/sence1.mp4 \
  --subtitle-input subtitles/sence1.srt \
  --output videos/sence1_subtitled.mp4
```

### 6. 合成影片與語音

使用 ffmpeg 將影片與語音音檔合成。

```sh
python app.py --task compose \
  --video-input videos/sence1.mp4 \
  --audio-input speaking-sound/sence1.mp3 \
  --output merged/sence1.mp4
```

### 7. 串接影片

可以逐一指定影片：

```sh
python app.py --task concat \
  --input merged/sence1.mp4 \
  --input merged/sence2.mp4 \
  --output final.mp4
```

也可以指定資料夾，CLI 會優先依 `sence*.mp4` 編號排序：

```sh
python app.py --task concat --input-dir merged --output final.mp4
```

## GUI

`gui.py` 是 CLI 的圖形化包裝：

- 支援模板、場景、語音、影片、字幕、影音合成、影片串接
- API Key 欄位預設讀取 `OPENAI_API_KEY_FOR_IMAGE` 與 `MINIMAX_API_KEY`
- 模板與場景的參考圖片欄位支援一次選多張，或以每行一個路徑手動貼上
- 語音聲線使用下拉選單選擇，不需手動輸入 voice ID
- 檔案欄位提供瀏覽按鈕
- 執行後會在下方顯示 stdout 與 stderr

GUI 不直接實作 API 邏輯，所有任務都會轉成 `app.py` CLI 指令執行。

## AI Skill 安裝

本專案已整理成可安裝的 AI coding-agent skill：

```text
skills/webpro-ppt-film-maker/SKILL.md
```

安裝時請複製整個 `skills/webpro-ppt-film-maker` 資料夾。Claude Code 可放到 `~/.claude/skills/`，opencode 可放到 `.opencode/skills/` 或 `~/.config/opencode/skills/`，其他支援 `SKILL.md` 的 Codex-like / OpenClaw-style 工具可放到其 skills 目錄，或參考 `skills/README.md`。

## JSON 輸出

所有任務都支援：

```sh
python app.py --task tts --text "hello" --output-json
```

錯誤時會輸出：

```text
執行失敗: 錯誤原因
```
