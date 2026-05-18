from __future__ import annotations

import argparse
import base64
import json
import math
import mimetypes
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from urllib import error, request
from typing import Any


OPENAI_API_BASE_URL = "https://api.openai.com/v1"
MINIMAX_API_BASE_URL = "https://api.minimax.io/v1"
OPENAI_IMAGE_API_KEY_ENV = "OPENAI_API_KEY_FOR_IMAGE"
LEGACY_OPENAI_API_KEY_ENV = "OPENAI_API_KEY"

DEFAULT_IMAGE_MODEL = "gpt-image-2"
DEFAULT_VIDEO_MODEL = "MiniMax-Hailuo-02"
DEFAULT_TTS_MODEL = "speech-2.8-hd"
DEFAULT_TTS_VOICE = "English_Graceful_Lady"
DEFAULT_AUDIO_FORMAT = "mp3"
DEFAULT_VIDEO_DURATION = 6
DEFAULT_VIDEO_RESOLUTION = "1080P"
DEFAULT_POLL_INTERVAL = 10
DEFAULT_TIMEOUT = 1800

SUPPORTED_REFERENCE_MIME_TYPES = {"image/jpeg", "image/png", "image/webp"}
MINIMAX_AUDIO_FORMATS = {"mp3", "wav", "opus", "flac", "pcm", "pcmu_raw", "pcmu_wav"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="PowerPoint 轉影片流程工具：模板、場景、TTS、影片、字幕、合成與串接。"
    )
    parser.add_argument(
        "--task",
        choices=["template", "scene", "tts", "video", "subtitle", "compose", "concat"],
        required=True,
        help="要執行的任務",
    )
    parser.add_argument("--prompt", help="圖片或影片生成提示詞")
    parser.add_argument(
        "--reference",
        action="append",
        default=[],
        help="OpenAI 圖片任務的參考圖片，可重複指定，template 建議 0 到 4 張",
    )
    parser.add_argument("--ppt-image", help="scene 任務使用的 PPT 匯出圖片")
    parser.add_argument("--template-image", help="scene 任務使用的模板圖片")
    parser.add_argument("--scene-number", type=int, help="scene 輸出編號，例如 1 會輸出 sence1.png")
    parser.add_argument("--image", help="video 任務使用的起始場景圖片")
    parser.add_argument("--end-image", help="video 任務使用的結束參考圖片，可選")
    parser.add_argument("--text", help="tts 任務的文稿內容，或字幕文字")
    parser.add_argument("--text-file", help="從文字檔讀取文稿或字幕")
    parser.add_argument("--video-input", help="subtitle/compose 任務使用的影片檔")
    parser.add_argument("--audio-input", help="video/compose 任務使用的語音音檔")
    parser.add_argument("--subtitle-input", help="subtitle 任務使用的字幕檔，例如 .srt")
    parser.add_argument("--input", action="append", default=[], help="concat 任務的輸入影片，可重複指定")
    parser.add_argument("--input-dir", help="concat 任務的輸入資料夾，會讀取 sence*.mp4 或 *.mp4")
    parser.add_argument("--output", help="輸出檔案路徑")
    parser.add_argument("--model", help="模型名稱；依任務有不同預設值")
    parser.add_argument("--size", default="1024x1024", help="OpenAI 圖片輸出尺寸，預設 1024x1024")
    parser.add_argument("--duration", type=int, help="影片秒數；未指定時預設 6，或可搭配 --audio-input 自動估算")
    parser.add_argument(
        "--resolution",
        default=DEFAULT_VIDEO_RESOLUTION,
        choices=["512P", "720P", "768P", "1080P", "512p", "720p", "768p", "1080p"],
        help="Hailuo 影片解析度，預設 1080P",
    )
    parser.add_argument("--voice", default=DEFAULT_TTS_VOICE, help="MiniMax TTS 聲線")
    parser.add_argument("--speed", type=float, default=1.0, help="MiniMax TTS 語速，範圍 0.5 到 2.0")
    parser.add_argument("--pitch", type=int, default=0, help="MiniMax TTS 音調，預設 0")
    parser.add_argument("--volume", type=float, default=1.0, help="MiniMax TTS 音量，預設 1.0")
    parser.add_argument(
        "--audio-format",
        default=DEFAULT_AUDIO_FORMAT,
        choices=sorted(MINIMAX_AUDIO_FORMATS),
        help="MiniMax TTS 音訊格式，預設 mp3",
    )
    parser.add_argument("--poll-interval", type=int, default=DEFAULT_POLL_INTERVAL, help="影片任務輪詢間隔秒數")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="影片任務等待完成秒數上限")
    parser.add_argument(
        "--openai-api-key",
        default=os.getenv(OPENAI_IMAGE_API_KEY_ENV) or os.getenv(LEGACY_OPENAI_API_KEY_ENV),
        help=f"OpenAI API Key；預設讀 {OPENAI_IMAGE_API_KEY_ENV}（相容 fallback: {LEGACY_OPENAI_API_KEY_ENV}）",
    )
    parser.add_argument("--minimax-api-key", default=os.getenv("MINIMAX_API_KEY"), help="MiniMax API Key；預設讀 MINIMAX_API_KEY")
    parser.add_argument("--api-key", help="相容欄位；會依 task 覆蓋對應 API Key")
    parser.add_argument("--output-json", action="store_true", help="輸出 JSON 結果")
    args = parser.parse_args()

    if args.api_key:
        if args.task in {"template", "scene"}:
            args.openai_api_key = args.api_key
        if args.task in {"tts", "video"}:
            args.minimax_api_key = args.api_key

    if args.task in {"template", "scene"}:
        if not args.openai_api_key:
            parser.error(
                f"此任務需要 --openai-api-key 或環境變數 {OPENAI_IMAGE_API_KEY_ENV}"
            )
        if not args.prompt:
            parser.error("此任務需要 --prompt")
        if not args.model:
            args.model = DEFAULT_IMAGE_MODEL
        if args.task == "template" and len(args.reference) > 4:
            parser.error("template 任務最多支援 4 張 --reference")
        if args.task == "scene" and not args.ppt_image:
            parser.error("scene 任務需要 --ppt-image")

    if args.task == "tts":
        if not args.minimax_api_key:
            parser.error("tts 任務需要 --minimax-api-key 或環境變數 MINIMAX_API_KEY")
        if not get_text_arg(args):
            parser.error("tts 任務需要 --text 或 --text-file")
        if not args.model:
            args.model = DEFAULT_TTS_MODEL
        if not 0.5 <= args.speed <= 2.0:
            parser.error("MiniMax TTS 的 --speed 必須介於 0.5 到 2.0")

    if args.task == "video":
        if not args.minimax_api_key:
            parser.error("video 任務需要 --minimax-api-key 或環境變數 MINIMAX_API_KEY")
        if not args.image:
            parser.error("video 任務需要 --image")
        if not args.prompt:
            parser.error("video 任務需要 --prompt")
        if not args.model:
            args.model = DEFAULT_VIDEO_MODEL
        args.resolution = args.resolution.upper()
        if args.end_image and args.model != "MiniMax-Hailuo-02":
            parser.error("起始加結束參考圖目前僅支援 --model MiniMax-Hailuo-02")
        if args.end_image and args.resolution == "512P":
            parser.error("使用 --end-image 時不支援 512P，請改用 768P 或 1080P")

    if args.task == "subtitle":
        if not args.video_input:
            parser.error("subtitle 任務需要 --video-input")
        if not args.subtitle_input:
            parser.error("subtitle 任務需要 --subtitle-input")

    if args.task == "compose":
        if not args.video_input:
            parser.error("compose 任務需要 --video-input")
        if not args.audio_input:
            parser.error("compose 任務需要 --audio-input")

    if args.task == "concat":
        if not args.input and not args.input_dir:
            parser.error("concat 任務需要至少一個 --input 或 --input-dir")

    return args


def get_text_arg(args: argparse.Namespace) -> str | None:
    if args.text_file:
        return Path(args.text_file).expanduser().read_text(encoding="utf-8").strip()
    if args.text:
        return args.text.strip()
    return None


def request_json(
    url: str,
    method: str,
    headers: dict[str, str],
    payload: dict[str, Any] | None = None,
    timeout: int = 120,
) -> dict[str, Any]:
    body = None
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
    req = request.Request(url=url, data=body, headers=headers, method=method)
    try:
        with request.urlopen(req, timeout=timeout) as response:
            raw_body = response.read().decode("utf-8")
            status_code = response.status
    except error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"API 呼叫失敗，HTTP {exc.code}: {error_body}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"無法連線到 API: {exc.reason}") from exc
    try:
        return json.loads(raw_body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"API 回傳非 JSON，HTTP {status_code}: {raw_body}") from exc


def request_binary(url: str, timeout: int = 600) -> bytes:
    req = request.Request(url=url, method="GET")
    try:
        with request.urlopen(req, timeout=timeout) as response:
            return response.read()
    except error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"下載失敗，HTTP {exc.code}: {error_body}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"下載時無法連線: {exc.reason}") from exc


def post_multipart_json(
    url: str,
    headers: dict[str, str],
    fields: dict[str, str],
    files: list[tuple[str, Path]],
    timeout: int = 300,
) -> dict[str, Any]:
    boundary = f"----webpro-{uuid.uuid4().hex}"
    body_parts: list[bytes] = []
    for name, value in fields.items():
        body_parts.append(f"--{boundary}\r\n".encode("utf-8"))
        body_parts.append(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("utf-8"))
        body_parts.append(str(value).encode("utf-8"))
        body_parts.append(b"\r\n")
    for field_name, path in files:
        mime_type, _ = mimetypes.guess_type(path.name)
        if mime_type is None:
            mime_type = "application/octet-stream"
        body_parts.append(f"--{boundary}\r\n".encode("utf-8"))
        body_parts.append(
            f'Content-Disposition: form-data; name="{field_name}"; filename="{path.name}"\r\n'.encode("utf-8")
        )
        body_parts.append(f"Content-Type: {mime_type}\r\n\r\n".encode("utf-8"))
        body_parts.append(path.read_bytes())
        body_parts.append(b"\r\n")
    body_parts.append(f"--{boundary}--\r\n".encode("utf-8"))
    body = b"".join(body_parts)

    multipart_headers = dict(headers)
    multipart_headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"
    multipart_headers["Content-Length"] = str(len(body))
    req = request.Request(url=url, data=body, headers=multipart_headers, method="POST")
    try:
        with request.urlopen(req, timeout=timeout) as response:
            raw_body = response.read().decode("utf-8")
            status_code = response.status
    except error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"API 呼叫失敗，HTTP {exc.code}: {error_body}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"無法連線到 API: {exc.reason}") from exc
    try:
        return json.loads(raw_body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"API 回傳非 JSON，HTTP {status_code}: {raw_body}") from exc


def ensure_existing_file(path_str: str, label: str) -> Path:
    path = Path(path_str).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"找不到{label}: {path}")
    return path


def ensure_command_available(command_name: str) -> None:
    if shutil.which(command_name) is None:
        raise RuntimeError(f"找不到系統指令: {command_name}")


def run_command(command: list[str]) -> None:
    try:
        subprocess.run(command, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as exc:
        details = exc.stderr.strip() or exc.stdout.strip() or "unknown error"
        raise RuntimeError(f"系統指令執行失敗: {' '.join(command)}: {details}") from exc


def run_command_capture(command: list[str]) -> str:
    try:
        result = subprocess.run(command, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as exc:
        details = exc.stderr.strip() or exc.stdout.strip() or "unknown error"
        raise RuntimeError(f"系統指令執行失敗: {' '.join(command)}: {details}") from exc
    return result.stdout.strip()


def get_media_duration_seconds(media_path: Path) -> float:
    ensure_command_available("ffprobe")
    output = run_command_capture([
        "ffprobe",
        "-v",
        "error",
        "-show_entries",
        "format=duration",
        "-of",
        "default=noprint_wrappers=1:nokey=1",
        str(media_path),
    ])
    try:
        duration = float(output)
    except ValueError as exc:
        raise RuntimeError(f"無法解析媒體長度: {output}") from exc
    if duration <= 0:
        raise RuntimeError(f"媒體長度無效: {duration}")
    return duration


def validate_image_reference(path: Path) -> None:
    mime_type, _ = mimetypes.guess_type(path.name)
    if mime_type not in SUPPORTED_REFERENCE_MIME_TYPES:
        raise ValueError(f"不支援的圖片格式: {path}，請使用 JPG、PNG 或 WebP")


def image_path_to_data_url(image_path: str) -> str:
    path = ensure_existing_file(image_path, "圖片")
    validate_image_reference(path)
    mime_type, _ = mimetypes.guess_type(path.name)
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def decode_openai_image_response(result: dict[str, Any], output_path: Path) -> None:
    if "error" in result:
        raise RuntimeError(json.dumps(result["error"], ensure_ascii=False))
    data = result.get("data")
    if not isinstance(data, list) or not data:
        raise RuntimeError(f"OpenAI 未回傳圖片資料: {json.dumps(result, ensure_ascii=False)}")
    first_image = data[0]
    if not isinstance(first_image, dict):
        raise RuntimeError(f"OpenAI 圖片資料格式異常: {json.dumps(result, ensure_ascii=False)}")
    if isinstance(first_image.get("b64_json"), str):
        output_path.write_bytes(base64.b64decode(first_image["b64_json"]))
        return
    if isinstance(first_image.get("url"), str):
        output_path.write_bytes(request_binary(first_image["url"]))
        return
    raise RuntimeError(f"OpenAI 圖片回傳缺少 b64_json 或 url: {json.dumps(result, ensure_ascii=False)}")


def openai_headers(api_key: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {api_key}"}


def generate_openai_image(
    api_key: str,
    model: str,
    prompt: str,
    size: str,
    output_path: Path,
    references: list[Path],
) -> dict[str, Any]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if references:
        for path in references:
            validate_image_reference(path)
        result = post_multipart_json(
            f"{OPENAI_API_BASE_URL}/images/edits",
            openai_headers(api_key),
            {
                "model": model,
                "prompt": prompt,
                "size": size,
            },
            [("image[]", path) for path in references],
        )
    else:
        result = request_json(
            f"{OPENAI_API_BASE_URL}/images/generations",
            "POST",
            {**openai_headers(api_key), "Content-Type": "application/json"},
            {
                "model": model,
                "prompt": prompt,
                "size": size,
                "n": 1,
            },
            timeout=300,
        )
    decode_openai_image_response(result, output_path)
    return result


def build_template_output_path(args: argparse.Namespace) -> Path:
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = Path.cwd() / "template.png"
    return output_path.with_suffix(".png") if output_path.suffix.lower() != ".png" else output_path


def build_scene_output_path(args: argparse.Namespace, ppt_path: Path) -> Path:
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        scene_number = args.scene_number
        if scene_number is None:
            digits = "".join(character for character in ppt_path.stem if character.isdigit())
            scene_number = int(digits) if digits else 1
        output_path = Path.cwd() / f"sence{scene_number}.png"
    return output_path.with_suffix(".png") if output_path.suffix.lower() != ".png" else output_path


def minimax_headers(api_key: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }


def ensure_minimax_success(result: dict[str, Any], action: str) -> dict[str, Any]:
    base_resp = result.get("base_resp")
    if not isinstance(base_resp, dict):
        raise RuntimeError(f"MiniMax {action} 回傳格式異常: {json.dumps(result, ensure_ascii=False)}")
    if base_resp.get("status_code") != 0:
        raise RuntimeError(f"MiniMax {action} 失敗: {json.dumps(result, ensure_ascii=False)}")
    return result


def build_minimax_tts_payload(args: argparse.Namespace, text: str) -> dict[str, Any]:
    return {
        "model": args.model,
        "text": text,
        "stream": False,
        "output_format": "hex",
        "voice_setting": {
            "voice_id": args.voice,
            "speed": args.speed,
            "vol": args.volume,
            "pitch": args.pitch,
        },
        "audio_setting": {
            "format": args.audio_format,
            "channel": 1,
        },
    }


def create_minimax_speech(api_key: str, payload: dict[str, Any]) -> tuple[bytes, str | None]:
    result = ensure_minimax_success(
        request_json(
            f"{MINIMAX_API_BASE_URL}/t2a_v2",
            "POST",
            minimax_headers(api_key),
            payload,
        ),
        "文字轉語音",
    )
    data = result.get("data")
    if not isinstance(data, dict):
        raise RuntimeError(f"MiniMax TTS 未回傳 data 欄位: {json.dumps(result, ensure_ascii=False)}")
    audio_hex = data.get("audio")
    if not isinstance(audio_hex, str) or not audio_hex:
        raise RuntimeError(f"MiniMax TTS 未回傳音訊資料: {json.dumps(result, ensure_ascii=False)}")
    try:
        audio_bytes = bytes.fromhex(audio_hex)
    except ValueError as exc:
        raise RuntimeError("MiniMax TTS 回傳的音訊資料不是有效 hex") from exc
    return audio_bytes, result.get("trace_id")


def build_audio_output_path(args: argparse.Namespace) -> Path:
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = Path.cwd() / f"speech_output.{args.audio_format}"
    if output_path.suffix.lower() != f".{args.audio_format}":
        output_path = output_path.with_suffix(f".{args.audio_format}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def infer_video_duration(args: argparse.Namespace) -> int:
    if args.duration:
        return args.duration
    if args.audio_input:
        audio_path = ensure_existing_file(args.audio_input, "音檔")
        return max(1, math.ceil(get_media_duration_seconds(audio_path)))
    return DEFAULT_VIDEO_DURATION


def build_video_payload(args: argparse.Namespace, first_frame_image: str, last_frame_image: str | None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": args.model,
        "prompt": args.prompt,
        "duration": infer_video_duration(args),
        "resolution": args.resolution,
        "first_frame_image": first_frame_image,
    }
    if last_frame_image is not None:
        payload["last_frame_image"] = last_frame_image
    return payload


def create_video_task(api_key: str, payload: dict[str, Any]) -> dict[str, Any]:
    return ensure_minimax_success(
        request_json(
            f"{MINIMAX_API_BASE_URL}/video_generation",
            "POST",
            minimax_headers(api_key),
            payload,
        ),
        "建立影片任務",
    )


def poll_video_task(api_key: str, task_id: str, poll_interval: int, timeout: int) -> dict[str, Any]:
    headers = minimax_headers(api_key)
    deadline = time.monotonic() + timeout
    last_result: dict[str, Any] | None = None
    while time.monotonic() < deadline:
        result = ensure_minimax_success(
            request_json(
                f"{MINIMAX_API_BASE_URL}/query/video_generation?task_id={task_id}",
                "GET",
                headers,
                timeout=120,
            ),
            "查詢影片任務",
        )
        last_result = result
        status = str(result.get("status", ""))
        if status == "Success":
            return result
        if status in {"Fail", "Failed", "Canceled", "Expired"}:
            raise RuntimeError(f"影片生成失敗，狀態為 {status}: {json.dumps(result, ensure_ascii=False)}")
        time.sleep(max(1, poll_interval))
    raise TimeoutError(
        f"等待影片完成逾時，最後狀態: {json.dumps(last_result, ensure_ascii=False) if last_result else 'unknown'}"
    )


def download_minimax_file(api_key: str, file_id: str, output_path: Path) -> None:
    headers = minimax_headers(api_key)
    headers.pop("Content-Type", None)
    retrieve_result = ensure_minimax_success(
        request_json(
            f"{MINIMAX_API_BASE_URL}/files/retrieve?file_id={file_id}",
            "GET",
            headers,
            timeout=120,
        ),
        "取得影片下載連結",
    )
    file_object = retrieve_result.get("file")
    if not isinstance(file_object, dict):
        raise RuntimeError(f"MiniMax 未回傳影片檔案資訊: {json.dumps(retrieve_result, ensure_ascii=False)}")
    download_url = file_object.get("download_url")
    if not isinstance(download_url, str) or not download_url:
        raise RuntimeError(f"MiniMax 未回傳下載網址: {json.dumps(retrieve_result, ensure_ascii=False)}")
    output_path.write_bytes(request_binary(download_url))


def build_video_output_path(args: argparse.Namespace, task_id: str) -> Path:
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        image_stem = Path(args.image).expanduser().stem or "hailuo_video"
        output_path = Path.cwd() / f"{image_stem}_{task_id}.mp4"
    if output_path.suffix.lower() != ".mp4":
        output_path = output_path.with_suffix(".mp4")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def escape_subtitles_filter_path(path: Path) -> str:
    value = str(path).replace("\\", "/")
    value = value.replace("'", "\\'").replace(":", "\\:")
    return f"'{value}'"


def build_mp4_output_path(args: argparse.Namespace, default_name: str) -> Path:
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = Path.cwd() / default_name
    if output_path.suffix.lower() != ".mp4":
        output_path = output_path.with_suffix(".mp4")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def burn_subtitle(video_path: Path, subtitle_path: Path, output_path: Path) -> None:
    ensure_command_available("ffmpeg")
    run_command([
        "ffmpeg",
        "-y",
        "-i",
        str(video_path),
        "-vf",
        f"subtitles={escape_subtitles_filter_path(subtitle_path)}",
        "-c:a",
        "copy",
        str(output_path),
    ])


def compose_video_audio(video_path: Path, audio_path: Path, output_path: Path) -> None:
    ensure_command_available("ffmpeg")
    run_command([
        "ffmpeg",
        "-y",
        "-i",
        str(video_path),
        "-i",
        str(audio_path),
        "-map",
        "0:v:0",
        "-map",
        "1:a:0",
        "-c:v",
        "copy",
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-shortest",
        str(output_path),
    ])


def collect_concat_inputs(args: argparse.Namespace) -> list[Path]:
    paths = [ensure_existing_file(path_str, "輸入影片") for path_str in args.input]
    if args.input_dir:
        input_dir = Path(args.input_dir).expanduser().resolve()
        if not input_dir.is_dir():
            raise FileNotFoundError(f"找不到輸入資料夾: {input_dir}")
        sence_paths = sorted(input_dir.glob("sence*.mp4"), key=scene_sort_key)
        paths.extend(sence_paths if sence_paths else sorted(input_dir.glob("*.mp4")))
    if not paths:
        raise RuntimeError("沒有找到可串接的影片")
    return paths


def scene_sort_key(path: Path) -> tuple[int, str]:
    digits = "".join(character for character in path.stem if character.isdigit())
    return (int(digits) if digits else 10**9, path.name)


def concat_videos(input_paths: list[Path], output_path: Path) -> None:
    ensure_command_available("ffmpeg")
    with tempfile.TemporaryDirectory(prefix="ppt-film-concat-") as temp_dir:
        list_path = Path(temp_dir) / "files.txt"
        lines = []
        for path in input_paths:
            safe_path = str(path).replace("'", "'\\''")
            lines.append(f"file '{safe_path}'")
        list_path.write_text("\n".join(lines), encoding="utf-8")
        run_command([
            "ffmpeg",
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(list_path),
            "-c",
            "copy",
            str(output_path),
        ])


def run_template(args: argparse.Namespace) -> dict[str, Any]:
    references = [ensure_existing_file(path_str, "參考圖片") for path_str in args.reference]
    output_path = build_template_output_path(args)
    api_result = generate_openai_image(
        args.openai_api_key,
        args.model,
        args.prompt,
        args.size,
        output_path,
        references,
    )
    return {"task": "template", "model": args.model, "output_path": str(output_path), "api_result": api_result}


def run_scene(args: argparse.Namespace) -> dict[str, Any]:
    ppt_path = ensure_existing_file(args.ppt_image, "PPT 圖片")
    references = [ppt_path]
    if args.template_image:
        references.append(ensure_existing_file(args.template_image, "模板圖片"))
    references.extend(ensure_existing_file(path_str, "參考圖片") for path_str in args.reference)
    output_path = build_scene_output_path(args, ppt_path)
    prompt = (
        "Use the PPT slide image as exact source content. Keep all slide text, layout, charts, colors, "
        "and readable details unchanged while fitting it into the requested template or scene. "
        f"User prompt: {args.prompt}"
    )
    api_result = generate_openai_image(args.openai_api_key, args.model, prompt, args.size, output_path, references)
    return {"task": "scene", "model": args.model, "output_path": str(output_path), "api_result": api_result}


def run_tts(args: argparse.Namespace) -> dict[str, Any]:
    text = get_text_arg(args)
    output_path = build_audio_output_path(args)
    payload = build_minimax_tts_payload(args, text)
    audio_bytes, trace_id = create_minimax_speech(args.minimax_api_key, payload)
    output_path.write_bytes(audio_bytes)
    return {
        "task": "tts",
        "model": args.model,
        "voice": args.voice,
        "audio_format": args.audio_format,
        "speed": args.speed,
        "trace_id": trace_id,
        "output_path": str(output_path),
    }


def run_video(args: argparse.Namespace) -> dict[str, Any]:
    first_frame_image = image_path_to_data_url(args.image)
    last_frame_image = image_path_to_data_url(args.end_image) if args.end_image else None
    payload = build_video_payload(args, first_frame_image, last_frame_image)
    if args.end_image and payload["resolution"] == "1080P" and payload["duration"] != 6:
        raise RuntimeError("MiniMax-Hailuo-02 使用 end-image 且解析度 1080P 時，目前僅支援 6 秒")
    create_result = create_video_task(args.minimax_api_key, payload)
    task_id = create_result.get("task_id")
    if not task_id:
        raise RuntimeError(f"MiniMax 未回傳 task_id: {json.dumps(create_result, ensure_ascii=False)}")
    final_result = poll_video_task(args.minimax_api_key, str(task_id), args.poll_interval, args.timeout)
    file_id = final_result.get("file_id")
    if not file_id:
        raise RuntimeError(f"MiniMax 任務成功但未回傳 file_id: {json.dumps(final_result, ensure_ascii=False)}")
    output_path = build_video_output_path(args, str(task_id))
    download_minimax_file(args.minimax_api_key, str(file_id), output_path)
    return {
        "task": "video",
        "model": args.model,
        "duration": payload["duration"],
        "task_id": task_id,
        "status": final_result.get("status"),
        "output_path": str(output_path),
        "create_result": create_result,
        "final_result": final_result,
    }


def run_subtitle(args: argparse.Namespace) -> dict[str, Any]:
    video_path = ensure_existing_file(args.video_input, "影片")
    subtitle_path = ensure_existing_file(args.subtitle_input, "字幕檔")
    output_path = build_mp4_output_path(args, f"{video_path.stem}_subtitled.mp4")
    burn_subtitle(video_path, subtitle_path, output_path)
    return {"task": "subtitle", "video_input": str(video_path), "subtitle_input": str(subtitle_path), "output_path": str(output_path)}


def run_compose(args: argparse.Namespace) -> dict[str, Any]:
    video_path = ensure_existing_file(args.video_input, "影片")
    audio_path = ensure_existing_file(args.audio_input, "音檔")
    output_path = build_mp4_output_path(args, f"{video_path.stem}_{audio_path.stem}_with_audio.mp4")
    compose_video_audio(video_path, audio_path, output_path)
    return {"task": "compose", "video_input": str(video_path), "audio_input": str(audio_path), "output_path": str(output_path)}


def run_concat(args: argparse.Namespace) -> dict[str, Any]:
    input_paths = collect_concat_inputs(args)
    output_path = build_mp4_output_path(args, "final.mp4")
    concat_videos(input_paths, output_path)
    return {"task": "concat", "inputs": [str(path) for path in input_paths], "output_path": str(output_path)}


def print_human_result(result: dict[str, Any]) -> None:
    task = result.get("task")
    if task == "template":
        print("模板圖片已完成")
        print(f"image: {result['output_path']}")
    elif task == "scene":
        print("場景圖片已完成")
        print(f"image: {result['output_path']}")
    elif task == "tts":
        print("MiniMax TTS 已完成")
        if result.get("trace_id"):
            print(f"trace_id: {result['trace_id']}")
        print(f"audio: {result['output_path']}")
    elif task == "video":
        print("Hailuo 影片已完成")
        print(f"task_id: {result['task_id']}")
        print(f"status: {result.get('status', 'unknown')}")
        print(f"video: {result['output_path']}")
    elif task == "subtitle":
        print("字幕影片已完成")
        print(f"video: {result['output_path']}")
    elif task == "compose":
        print("影片與語音合成已完成")
        print(f"video: {result['output_path']}")
    elif task == "concat":
        print("影片串接已完成")
        print(f"count: {len(result['inputs'])}")
        print(f"video: {result['output_path']}")


def main() -> int:
    args = parse_args()
    runners = {
        "template": run_template,
        "scene": run_scene,
        "tts": run_tts,
        "video": run_video,
        "subtitle": run_subtitle,
        "compose": run_compose,
        "concat": run_concat,
    }
    try:
        result = runners[args.task](args)
    except Exception as exc:
        print(f"執行失敗: {exc}", file=sys.stderr)
        return 1
    if args.output_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_human_result(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
