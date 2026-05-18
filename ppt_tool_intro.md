# PowerPoint 轉影片工具

## 投影片 1：PowerPoint 轉影片工具

把簡報內容快速轉成可播放的影片。

- 支援從 PPT 投影片圖片開始製作
- 可產生模板、場景圖、旁白、影片片段
- 最後合成完整影片

## 投影片 2：這個工具解決什麼問題？

傳統做法需要手動剪輯、配音、合成。

- PPT 內容已經準備好，但還不是影片
- 每頁投影片需要變成有畫面動態的場景
- 旁白、字幕、影片合成流程繁瑣
- 多個影片片段需要依序串接

這個工具把流程拆成可重複執行的任務。

## 投影片 3：工具提供兩種入口

- `app.py`：CLI 指令工具，適合自動化與批次處理
- `gui.py`：圖形化介面，適合手動操作與測試

GUI 不直接處理 API 或影片邏輯，而是把使用者輸入轉成 CLI 指令執行。

## 投影片 4：核心能力

- 產生模板畫面
- 將 PPT 圖片合成為場景圖片
- 將旁白文稿轉成語音
- 用場景圖片產生影片片段
- 燒入字幕
- 合成影片與語音
- 串接多段影片成最終成品

## 投影片 5：需要的外部服務與工具

- OpenAI Images API：產生模板與場景圖片
- MiniMax TTS API：將文稿轉成旁白語音
- MiniMax Hailuo API：用圖片產生影片片段
- ffmpeg / ffprobe：處理字幕、影音合成與片段串接

API Key 透過環境變數設定：

- `OPENAI_API_KEY_FOR_IMAGE`
- `MINIMAX_API_KEY`

## 投影片 6：從 PPT 變影片的整體流程

1. 製作 PowerPoint 簡報
2. 將每頁 PPT 匯出成圖片
3. 產生影片用的模板畫面
4. 把每頁 PPT 圖片放入模板，產生場景圖片
5. 將旁白文稿轉成語音檔
6. 依照語音長度產生對應影片片段
7. 合成影片片段與語音
8. 將所有片段串接成完整影片

## 投影片 7：第一步，準備 PPT 與模板

先完成簡報內容，並將每一頁匯出為 PNG 或 JPEG。

接著使用 `template` 任務產生影片共用模板，例如：

```sh
python app.py --task template \
  --prompt "高質感新聞攝影棚，大型簡報牆" \
  --output template.png
```

模板會決定影片的整體視覺風格。

## 投影片 8：第二步，產生每一頁的場景圖

使用 `scene` 任務，將 PPT 匯出的圖片放進模板中。

```sh
python app.py --task scene \
  --ppt-image slide1.png \
  --template-image template.png \
  --scene-number 1 \
  --prompt "保留 PPT 文字清晰可讀" \
  --output sence1.png
```

每一頁會產生一張對應的 `sence*.png` 場景圖片。

## 投影片 9：第三步，產生旁白與影片片段

使用 `tts` 將文稿轉成語音。

```sh
python app.py --task tts \
  --text-file script1.txt \
  --output speaking-sound/sence1.mp3
```

再使用 `video` 讓場景圖變成影片片段。

```sh
python app.py --task video \
  --image sence1.png \
  --audio-input speaking-sound/sence1.mp3 \
  --prompt "鏡頭緩慢推進，畫面自然流動" \
  --output videos/sence1.mp4
```

## 投影片 10：第四步，合成與串接

將影片畫面與旁白語音合成：

```sh
python app.py --task compose \
  --video-input videos/sence1.mp4 \
  --audio-input speaking-sound/sence1.mp3 \
  --output merged/sence1.mp4
```

最後串接所有片段：

```sh
python app.py --task concat \
  --input-dir merged \
  --output final.mp4
```

## 投影片 11：GUI 適合怎麼用？

GUI 適合不熟悉指令的人操作。

- 每個任務都有獨立頁籤
- 可用檔案選擇器選圖片、語音、字幕與輸出位置
- API Key 可直接填入欄位
- 執行結果會顯示 stdout 與 stderr，方便除錯

啟動方式：

```sh
python gui.py
```

## 投影片 12：總結

這個工具把 PPT 轉影片拆成清楚的流水線。

- PPT 負責內容
- OpenAI 負責模板與場景視覺
- MiniMax 負責旁白與影片生成
- ffmpeg 負責媒體合成
- CLI 與 GUI 讓流程可以手動操作，也可以自動化

最終目標：用可重複、可調整的流程，把簡報快速變成完整影片。
