from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


OUT = Path("ppt_tool_intro.pptx")
TOTAL_W = 13_333_500
TOTAL_H = 7_500_000

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

FONT = "Microsoft JhengHei"


slides = [
    {
        "kind": "cover",
        "kicker": "AI WORKFLOW TOOL",
        "title": "PowerPoint 轉影片工具",
        "subtitle": "把簡報內容轉成有旁白、有畫面動態、可直接播放的完整影片",
        "points": ["CLI 批次處理", "GUI 手動操作", "OpenAI + MiniMax + ffmpeg"],
        "accent": "2563EB",
    },
    {
        "kind": "problem",
        "title": "這個工具解決什麼問題？",
        "subtitle": "把原本分散的剪輯、配音、合成工作，整理成可重複執行的流程。",
        "cards": [
            ("內容已完成", "PPT 已經有邏輯與文字，但還不是影片格式"),
            ("製作太分散", "圖片、旁白、字幕、剪輯常需要不同工具處理"),
            ("難以重複", "每次修改都要重新人工整理片段與順序"),
        ],
        "accent": "0F766E",
    },
    {
        "kind": "split",
        "title": "兩種使用入口",
        "subtitle": "同一套核心流程，可以用指令自動化，也可以用 GUI 操作。",
        "left_title": "app.py",
        "left_body": ["CLI 指令工具", "適合批次處理與自動化", "所有 API 與媒體處理邏輯都在這裡"],
        "right_title": "gui.py",
        "right_body": ["tkinter 圖形化介面", "適合手動操作與快速測試", "把使用者輸入轉成 CLI 指令執行"],
        "accent": "7C3AED",
    },
    {
        "kind": "cards",
        "title": "核心能力",
        "subtitle": "每一個能力都對應 `app.py --task` 的一個任務。",
        "items": [
            ("template", "產生模板畫面"),
            ("scene", "PPT 圖片變場景圖"),
            ("tts", "文稿轉旁白語音"),
            ("video", "圖片生成影片片段"),
            ("subtitle", "燒入字幕"),
            ("compose", "合成影音"),
            ("concat", "串接最終影片"),
        ],
        "accent": "EA580C",
    },
    {
        "kind": "ecosystem",
        "title": "需要的外部服務與工具",
        "subtitle": "這個工具負責編排流程，外部服務負責生成與媒體處理。",
        "items": [
            ("OpenAI Images API", "產生模板與場景圖片"),
            ("MiniMax TTS API", "將文稿轉成旁白語音"),
            ("MiniMax Hailuo API", "用場景圖片生成影片片段"),
            ("ffmpeg / ffprobe", "字幕、影音合成、片段串接與長度偵測"),
        ],
        "note": "API Key：OPENAI_API_KEY_FOR_IMAGE、MINIMAX_API_KEY",
        "accent": "DC2626",
    },
    {
        "kind": "timeline",
        "title": "從 PPT 變影片的整體流程",
        "subtitle": "流程不是一次完成，而是分成可檢查、可重跑的中間成果。",
        "steps": [
            "PPT",
            "匯出圖片",
            "模板",
            "場景圖",
            "旁白",
            "影片片段",
            "影音合成",
            "Final MP4",
        ],
        "accent": "0891B2",
    },
    {
        "kind": "code",
        "title": "第一步：準備 PPT 與模板",
        "subtitle": "先完成簡報內容，將每頁匯出成 PNG 或 JPEG，再建立共用視覺模板。",
        "body": ["模板會決定影片整體風格，例如新聞棚、教學場景、企業簡報牆。"],
        "code": "python app.py --task template --prompt \"高質感新聞攝影棚，大型簡報牆\" --output template.png",
        "accent": "4F46E5",
    },
    {
        "kind": "code",
        "title": "第二步：每一頁 PPT 產生場景圖",
        "subtitle": "把 PPT 匯出圖片放進模板，產生每一頁對應的影片場景。",
        "body": ["輸出檔依照目前專案習慣命名為 sence1.png、sence2.png、sence3.png。"],
        "code": "python app.py --task scene --ppt-image slide1.png --template-image template.png --scene-number 1 --output sence1.png",
        "accent": "16A34A",
    },
    {
        "kind": "code",
        "title": "第三步：產生旁白與影片片段",
        "subtitle": "先用 TTS 產生旁白，再用場景圖生成影片片段。",
        "body": ["如果 video 任務提供 audio-input，工具可以用 ffprobe 依音檔長度估算影片秒數。"],
        "code": "python app.py --task tts --text-file script1.txt --output speaking-sound/sence1.mp3\npython app.py --task video --image sence1.png --audio-input speaking-sound/sence1.mp3 --output videos/sence1.mp4",
        "accent": "9333EA",
    },
    {
        "kind": "code",
        "title": "第四步：合成與串接",
        "subtitle": "每一段影片先和旁白合成，最後把所有片段串成完整影片。",
        "body": ["concat 任務指定資料夾時，會優先依 sence*.mp4 編號排序。"],
        "code": "python app.py --task compose --video-input videos/sence1.mp4 --audio-input speaking-sound/sence1.mp3 --output merged/sence1.mp4\npython app.py --task concat --input-dir merged --output final.mp4",
        "accent": "C2410C",
    },
    {
        "kind": "gui",
        "title": "GUI 適合怎麼用？",
        "subtitle": "不熟悉指令時，可以用 GUI 逐步操作同一套流程。",
        "items": ["每個任務都有獨立頁籤", "可用檔案選擇器選圖片、語音、字幕與輸出位置", "API Key 可直接填入欄位", "下方顯示 stdout 與 stderr，方便除錯"],
        "code": "python gui.py",
        "accent": "0D9488",
    },
    {
        "kind": "summary",
        "title": "總結",
        "subtitle": "用清楚的流水線，把 PPT 變成完整影片。",
        "items": [
            ("PPT", "負責內容與邏輯"),
            ("OpenAI", "負責模板與場景視覺"),
            ("MiniMax", "負責旁白與影片生成"),
            ("ffmpeg", "負責字幕、影音合成與串接"),
            ("CLI / GUI", "讓流程可以自動化，也可以手動操作"),
        ],
        "accent": "1D4ED8",
    },
]


def xml(text: str) -> str:
    return escape(text, {"\"": "&quot;"})


def run(text: str, size: int, color: str = "0F172A", bold: bool = False, font: str = FONT) -> str:
    bold_xml = "<a:b/>" if bold else ""
    return (
        f'<a:r><a:rPr lang="zh-TW" sz="{size}">{bold_xml}'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>'
        f'</a:rPr><a:t>{xml(text)}</a:t></a:r>'
    )


def paragraph(text: str, size: int, color: str = "0F172A", bold: bool = False, bullet: bool = False, center: bool = False, font: str = FONT) -> str:
    algn = ' algn="ctr"' if center else ""
    if bullet:
        ppr = f'<a:pPr marL="380000" indent="-220000"{algn}><a:buChar char="•"/></a:pPr>'
    else:
        ppr = f'<a:pPr{algn}><a:buNone/></a:pPr>'
    return f'<a:p>{ppr}{run(text, size, color, bold, font)}<a:endParaRPr lang="zh-TW" sz="{size}"/></a:p>'


def text_box(id_: int, x: int, y: int, cx: int, cy: int, lines: list[str], size: int = 2200, color: str = "0F172A", bold: bool = False, bullet: bool = False, center: bool = False, font: str = FONT) -> str:
    body = "".join(paragraph(line, size, color, bold, bullet, center, font) for line in lines)
    return f'''
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{id_}" name="Text {id_}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
      <p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>{body}</p:txBody>
    </p:sp>'''


def rect(id_: int, x: int, y: int, cx: int, cy: int, fill: str, radius: bool = False, line: str | None = None, alpha: int | None = None) -> str:
    alpha_xml = f'<a:alpha val="{alpha}"/>' if alpha else ""
    line_xml = f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    prst = "roundRect" if radius else "rect"
    return f'''
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{id_}" name="Shape {id_}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}">{alpha_xml}</a:srgbClr></a:solidFill>{line_xml}</p:spPr>
      <p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>
    </p:sp>'''


def title_block(shapes: list[str], slide: dict[str, object], start_id: int = 10) -> int:
    accent = str(slide["accent"])
    shapes.append(text_box(start_id, 720000, 520000, 10_900_000, 520000, [str(slide["title"])], 3300, "0F172A", True))
    shapes.append(rect(start_id + 1, 720000, 1_130_000, 920000, 70000, accent, True))
    subtitle = slide.get("subtitle")
    if subtitle:
        shapes.append(text_box(start_id + 2, 720000, 1_280_000, 11_200_000, 430000, [str(subtitle)], 1700, "475569"))
    return start_id + 3


def base_shapes(accent: str, index: int) -> list[str]:
    return [
        rect(2, 0, 0, TOTAL_W, TOTAL_H, "F8FAFC"),
        rect(3, 0, 0, TOTAL_W, 180000, accent),
        rect(4, 12_760_000, 180000, 280000, 6_980_000, accent, alpha=15000),
        text_box(5, 720000, 6_910_000, 3_000_000, 230000, ["webpro_ppt_film_maker"], 1050, "64748B"),
        text_box(6, 11_800_000, 6_910_000, 720000, 230000, [f"{index:02d}"], 1150, "64748B", center=True),
    ]


def card(shapes: list[str], id_: int, x: int, y: int, cx: int, cy: int, title: str, body: str, accent: str) -> int:
    shapes.append(rect(id_, x, y, cx, cy, "FFFFFF", True, "E2E8F0"))
    shapes.append(rect(id_ + 1, x, y, 95000, cy, accent, True))
    shapes.append(text_box(id_ + 2, x + 260000, y + 220000, cx - 420000, 300000, [title], 1900, "0F172A", True))
    shapes.append(text_box(id_ + 3, x + 260000, y + 620000, cx - 420000, cy - 760000, [body], 1450, "475569"))
    return id_ + 4


def render_slide(slide: dict[str, object], index: int) -> str:
    accent = str(slide["accent"])
    shapes = base_shapes(accent, index)
    kind = slide["kind"]

    if kind == "cover":
        shapes.append(rect(10, 0, 0, TOTAL_W, TOTAL_H, "0F172A"))
        shapes.append(rect(11, 0, 0, 4_200_000, TOTAL_H, accent, alpha=88000))
        shapes.append(rect(12, 800000, 1_030_000, 1_850_000, 90000, "93C5FD", True))
        shapes.append(text_box(13, 800000, 1_220_000, 3_800_000, 280000, [str(slide["kicker"])], 1100, "BFDBFE", True))
        shapes.append(text_box(14, 800000, 1_730_000, 8_500_000, 900000, [str(slide["title"])], 4100, "FFFFFF", True))
        shapes.append(text_box(15, 820000, 2_740_000, 8_900_000, 600000, [str(slide["subtitle"])], 1900, "CBD5E1"))
        x = 850000
        sid = 16
        for point in slide["points"]:
            shapes.append(rect(sid, x, 4_280_000, 2_550_000, 620000, "FFFFFF", True, alpha=14000))
            shapes.append(text_box(sid + 1, x + 180000, 4_450_000, 2_180_000, 260000, [point], 1400, "FFFFFF", True, center=True))
            x += 2_800_000
            sid += 2
        shapes.append(text_box(30, 800000, 6_820_000, 4_500_000, 250000, ["OpenAI Images | MiniMax TTS / Hailuo | ffmpeg"], 1050, "94A3B8"))

    elif kind == "problem":
        next_id = title_block(shapes, slide)
        x = 820000
        for title, body in slide["cards"]:
            next_id = card(shapes, next_id, x, 2_200_000, 3_650_000, 2_600_000, title, body, accent)
            x += 4_020_000
        shapes.append(rect(next_id, 1_300_000, 5_400_000, 10_700_000, 780000, "ECFDF5", True, "BBF7D0"))
        shapes.append(text_box(next_id + 1, 1_600_000, 5_610_000, 10_100_000, 280000, ["解法：把 PPT 轉影片拆成模板、場景圖、旁白、影片片段、合成與串接等可重複任務。"], 1650, "065F46", True, center=True))

    elif kind == "split":
        next_id = title_block(shapes, slide)
        shapes.append(rect(next_id, 900000, 2_100_000, 5_380_000, 3_920_000, "FFFFFF", True, "DDD6FE"))
        shapes.append(rect(next_id + 1, 7_050_000, 2_100_000, 5_380_000, 3_920_000, "FFFFFF", True, "DDD6FE"))
        shapes.append(text_box(next_id + 2, 1_250_000, 2_460_000, 4_600_000, 420000, [str(slide["left_title"])], 2600, accent, True, center=True))
        shapes.append(text_box(next_id + 3, 7_400_000, 2_460_000, 4_600_000, 420000, [str(slide["right_title"])], 2600, accent, True, center=True))
        shapes.append(text_box(next_id + 4, 1_380_000, 3_230_000, 4_250_000, 1_950_000, list(slide["left_body"]), 1650, "334155", bullet=True))
        shapes.append(text_box(next_id + 5, 7_530_000, 3_230_000, 4_250_000, 1_950_000, list(slide["right_body"]), 1650, "334155", bullet=True))

    elif kind == "cards":
        next_id = title_block(shapes, slide)
        x0, y0 = 800000, 2_040_000
        for i, (task, desc) in enumerate(slide["items"]):
            col = i % 4
            row = i // 4
            x = x0 + col * 3_050_000
            y = y0 + row * 1_680_000
            shapes.append(rect(next_id, x, y, 2_700_000, 1_250_000, "FFFFFF", True, "E2E8F0"))
            shapes.append(text_box(next_id + 1, x + 220000, y + 230000, 2_250_000, 260000, [task], 1600, accent, True))
            shapes.append(text_box(next_id + 2, x + 220000, y + 620000, 2_250_000, 300000, [desc], 1280, "475569"))
            next_id += 3

    elif kind == "ecosystem":
        next_id = title_block(shapes, slide)
        y = 2_050_000
        for name, body in slide["items"]:
            shapes.append(rect(next_id, 1_000_000, y, 11_300_000, 790000, "FFFFFF", True, "E2E8F0"))
            shapes.append(text_box(next_id + 1, 1_350_000, y + 195000, 3_000_000, 260000, [name], 1500, accent, True))
            shapes.append(text_box(next_id + 2, 4_500_000, y + 195000, 7_100_000, 260000, [body], 1450, "334155"))
            y += 980000
            next_id += 3
        shapes.append(rect(next_id, 1_000_000, 6_150_000, 11_300_000, 520000, "FEF2F2", True, "FECACA"))
        shapes.append(text_box(next_id + 1, 1_250_000, 6_295_000, 10_800_000, 200000, [str(slide["note"])], 1250, "991B1B", True, center=True))

    elif kind == "timeline":
        next_id = title_block(shapes, slide)
        x, y = 700000, 2_450_000
        for i, step in enumerate(slide["steps"]):
            shapes.append(rect(next_id, x, y, 1_340_000, 900000, "FFFFFF", True, "BAE6FD"))
            shapes.append(text_box(next_id + 1, x + 140000, y + 280000, 1_060_000, 220000, [step], 1120, "0F172A", True, center=True))
            if i < len(slide["steps"]) - 1:
                shapes.append(text_box(next_id + 2, x + 1_420_000, y + 280000, 240000, 220000, ["→"], 1600, accent, True, center=True))
            x += 1_570_000
            next_id += 3
        shapes.append(rect(next_id, 1_050_000, 4_450_000, 11_200_000, 1_050_000, "ECFEFF", True, "A5F3FC"))
        shapes.append(text_box(next_id + 1, 1_450_000, 4_720_000, 10_400_000, 340000, ["每一步都有中間檔案，因此可以單獨檢查、單獨重跑，也方便定位問題。"], 1700, "155E75", True, center=True))

    elif kind == "code":
        next_id = title_block(shapes, slide)
        shapes.append(rect(next_id, 850000, 2_140_000, 11_700_000, 1_150_000, "FFFFFF", True, "E2E8F0"))
        shapes.append(text_box(next_id + 1, 1_180_000, 2_420_000, 10_900_000, 420000, list(slide["body"]), 1600, "334155", bullet=True))
        shapes.append(rect(next_id + 2, 850000, 3_650_000, 11_700_000, 2_050_000, "111827", True))
        shapes.append(rect(next_id + 3, 850000, 3_650_000, 11_700_000, 360000, accent, True, alpha=88000))
        shapes.append(text_box(next_id + 4, 1_120_000, 3_725_000, 4_000_000, 160000, ["COMMAND"], 850, "E0F2FE", True))
        code_lines = str(slide["code"]).split("\n")
        shapes.append(text_box(next_id + 5, 1_100_000, 4_230_000, 10_900_000, 1_050_000, code_lines, 980, "E5E7EB", font="Menlo"))

    elif kind == "gui":
        next_id = title_block(shapes, slide)
        shapes.append(rect(next_id, 900000, 2_070_000, 6_950_000, 3_780_000, "FFFFFF", True, "CCFBF1"))
        shapes.append(text_box(next_id + 1, 1_250_000, 2_440_000, 6_100_000, 1_900_000, list(slide["items"]), 1600, "334155", bullet=True))
        shapes.append(rect(next_id + 2, 8_250_000, 2_070_000, 3_850_000, 1_230_000, "0F172A", True))
        shapes.append(text_box(next_id + 3, 8_550_000, 2_520_000, 3_250_000, 240000, [str(slide["code"])], 1150, "E5E7EB", True, center=True, font="Menlo"))
        shapes.append(rect(next_id + 4, 8_250_000, 3_650_000, 3_850_000, 2_200_000, "F0FDFA", True, "99F6E4"))
        shapes.append(text_box(next_id + 5, 8_560_000, 4_130_000, 3_250_000, 850000, ["GUI 是同一套 CLI 的操作面板，不重複實作 API 邏輯。"], 1550, "115E59", True, center=True))

    elif kind == "summary":
        next_id = title_block(shapes, slide)
        x, y = 950000, 2_050_000
        for i, (name, body) in enumerate(slide["items"]):
            shapes.append(rect(next_id, x, y, 2_200_000, 1_550_000, "FFFFFF", True, "BFDBFE"))
            shapes.append(text_box(next_id + 1, x + 200000, y + 240000, 1_800_000, 260000, [name], 1550, accent, True, center=True))
            shapes.append(text_box(next_id + 2, x + 220000, y + 700000, 1_760_000, 420000, [body], 1180, "475569", center=True))
            x += 2_420_000
            next_id += 3
        shapes.append(rect(next_id, 1_300_000, 5_330_000, 10_700_000, 760000, accent, True))
        shapes.append(text_box(next_id + 1, 1_650_000, 5_555_000, 10_000_000, 250000, ["目標：用可重複、可調整的流程，把簡報快速變成完整影片。"], 1650, "FFFFFF", True, center=True))

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    {''.join(shapes)}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


def content_types() -> str:
    overrides = [
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    overrides += [f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(1, len(slides) + 1)]
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {''.join(overrides)}
</Types>'''


def root_rels() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def presentation() -> str:
    slide_ids = "".join(f'<p:sldId id="{256 + i}" r:id="rId{i + 1}"/>' for i in range(1, len(slides) + 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}" saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="{TOTAL_W}" cy="{TOTAL_H}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>'''


def presentation_rels() -> str:
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    rels += [f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>' for i in range(1, len(slides) + 1)]
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{''.join(rels)}</Relationships>'''


def slide_rels() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''


def slide_master() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>'''


def slide_master_rels() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>'''


def slide_layout() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>'''


def slide_layout_rels() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''


def theme() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="{NS_A}" name="Webpro Film Maker">
  <a:themeElements>
    <a:clrScheme name="Webpro"><a:dk1><a:srgbClr val="0F172A"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="334155"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="0F766E"/></a:accent2><a:accent3><a:srgbClr val="7C3AED"/></a:accent3><a:accent4><a:srgbClr val="EA580C"/></a:accent4><a:accent5><a:srgbClr val="0891B2"/></a:accent5><a:accent6><a:srgbClr val="DC2626"/></a:accent6><a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Webpro"><a:majorFont><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:majorFont><a:minorFont><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Webpro"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>'''


def core_props() -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>PowerPoint 轉影片工具</dc:title><dc:subject>工具介紹與 PPT 轉影片流程</dc:subject><dc:creator>OpenCode</dc:creator><cp:lastModifiedBy>OpenCode</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def app_props() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>OpenCode</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>{len(slides)}</Slides></Properties>'''


def build() -> None:
    with ZipFile(OUT, "w", ZIP_DEFLATED) as pptx:
        pptx.writestr("[Content_Types].xml", content_types())
        pptx.writestr("_rels/.rels", root_rels())
        pptx.writestr("docProps/core.xml", core_props())
        pptx.writestr("docProps/app.xml", app_props())
        pptx.writestr("ppt/presentation.xml", presentation())
        pptx.writestr("ppt/_rels/presentation.xml.rels", presentation_rels())
        pptx.writestr("ppt/slideMasters/slideMaster1.xml", slide_master())
        pptx.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels())
        pptx.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout())
        pptx.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels())
        pptx.writestr("ppt/theme/theme1.xml", theme())
        for i, slide in enumerate(slides, start=1):
            pptx.writestr(f"ppt/slides/slide{i}.xml", render_slide(slide, i))
            pptx.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide_rels())


if __name__ == "__main__":
    build()
    print(OUT)
