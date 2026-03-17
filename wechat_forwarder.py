import argparse
import ctypes
import json
import logging
import re
import sys
import time
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable

import uiautomation as auto
import cv2
import imagehash
import numpy as np
from PIL import ImageChops, ImageGrab
from rapidocr_onnxruntime import RapidOCR


@dataclass
class Config:
    source_chat: str
    target_chats: list[str]
    wechat_window_title: str = "微信"
    settle_delay_seconds: float = 1.2
    dialog_delay_seconds: float = 1.0
    max_message_scan: int = 60
    search_box_x_ratio: float = 0.20
    search_box_y_ratio: float = 0.10
    text_message_click_x_ratio: float = 0.68
    text_message_click_y_ratio: float = 0.10
    message_click_x_ratio: float = 0.78
    message_click_y_ratio: float = 0.62
    grouped_video_click_x_ratio: float = 0.79
    grouped_video_click_y_ratio: float = 0.33
    multi_select_checkbox_x_ratio: float = 0.43
    text_forward_menu_y_offset: int = 175
    video_forward_menu_y_offset: int = 26
    multi_select_menu_y_offset: int = 283
    multi_select_menu_row_index: int = 7
    multi_select_menu_click_ratio: float = 0.67
    multi_forward_button_x_ratio: float = 0.50
    multi_forward_button_y_ratio: float = 0.85
    forward_dialog_search_x_ratio: float = 0.19
    forward_dialog_search_y_ratio: float = 0.07
    forward_dialog_clear_x_ratio: float = 0.42
    forward_dialog_result_x_ratio: float = 0.20
    forward_dialog_result_y_ratio: float = 0.25
    forward_dialog_send_x_ratio: float = 0.60
    forward_dialog_send_y_ratio: float = 0.88
    forward_dialog_cancel_x_ratio: float = 0.86
    forward_dialog_cancel_y_ratio: float = 0.87
    forward_text_then_video: bool = True
    forward_as_grouped_pair: bool = True
    batch_mode_enabled: bool = True
    max_pairs_per_run: int = 30
    max_scroll_pages: int = 40
    stop_after_seen_pairs: int = 2
    stop_after_empty_pages: int = 3
    scroll_wheel_delta: int = 1440
    state_file: str = "forward_state.json"
    dry_run: bool = True
    preview_forward_dialog: bool = False


@dataclass
class PairCandidate:
    text_x_ratio: float
    text_y_ratio: float
    video_x_ratio: float
    video_y_ratio: float
    text_menu_offset: int
    signature: str
    text_content: str
    approx_top_ratio: float


OCR_ENGINE = RapidOCR()
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def load_config(path: Path) -> Config:
    raw = json.loads(path.read_text(encoding="utf-8"))
    return Config(**raw)


def wait(seconds: float) -> None:
    time.sleep(seconds)


def window(title: str) -> auto.WindowControl:
    win = auto.WindowControl(searchDepth=1, Name=title)
    if not win.Exists(3):
        raise RuntimeError(f"未找到微信窗口: {title}")
    return win


def ensure_wechat_window(title: str, timeout_seconds: float = 8) -> auto.WindowControl:
    deadline = time.time() + timeout_seconds
    last_error: Exception | None = None
    while time.time() < deadline:
        try:
            return window(title)
        except Exception as exc:
            last_error = exc
            try:
                ctypes.windll.shell32.ShellExecuteW(
                    None,
                    "open",
                    r"D:\Program Files\Tencent\Weixin\Weixin.exe",
                    None,
                    None,
                    1,
                )
            except Exception:
                pass
            wait(0.5)
    raise RuntimeError(str(last_error) if last_error else f"未找到微信窗口: {title}")


def safe_window_rect(win: auto.WindowControl, title: str) -> auto.Rect:
    try:
        return win.BoundingRectangle
    except Exception:
        refreshed = window(title)
        return refreshed.BoundingRectangle


def activate_wechat(win: auto.WindowControl) -> None:
    win.SetActive()
    try:
        win.SwitchToThisWindow()
    except Exception:
        pass


def click_at(x: int, y: int, right: bool = False) -> None:
    ctypes.windll.user32.SetCursorPos(x, y)
    wait(0.1)
    if right:
        ctypes.windll.user32.mouse_event(0x0008, 0, 0, 0, 0)
        ctypes.windll.user32.mouse_event(0x0010, 0, 0, 0, 0)
    else:
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)


def press_key(vk_code: int) -> None:
    ctypes.windll.user32.keybd_event(vk_code, 0, 0, 0)
    wait(0.05)
    ctypes.windll.user32.keybd_event(vk_code, 0, 0x0002, 0)


def press_ctrl_shortcut(vk_code: int) -> None:
    ctypes.windll.user32.keybd_event(0x11, 0, 0, 0)
    wait(0.05)
    press_key(vk_code)
    wait(0.05)
    ctypes.windll.user32.keybd_event(0x11, 0, 0x0002, 0)


def clear_active_input(repeats: int = 30) -> None:
    for _ in range(repeats):
        press_key(0x08)
    for _ in range(repeats):
        press_key(0x2E)


def mouse_wheel(delta: int) -> None:
    ctypes.windll.user32.mouse_event(0x0800, 0, 0, delta, 0)


def first_existing(controls: Iterable[auto.Control]) -> auto.Control | None:
    for control in controls:
        try:
            if control.Exists(0.5):
                return control
        except Exception:
            continue
    return None


def set_clipboard_text(text: str) -> None:
    data = ctypes.create_unicode_buffer(text)
    size = ctypes.sizeof(data)
    handle = ctypes.c_void_p(ctypes.windll.kernel32.GlobalAlloc(GMEM_MOVEABLE, size))
    if not handle.value:
        raise OSError("GlobalAlloc failed")

    locked = ctypes.c_void_p(ctypes.windll.kernel32.GlobalLock(handle))
    if not locked:
        ctypes.windll.kernel32.GlobalFree(handle)
        raise OSError("GlobalLock failed")

    ctypes.memmove(locked, ctypes.addressof(data), size)
    ctypes.windll.kernel32.GlobalUnlock(handle)

    if not ctypes.windll.user32.OpenClipboard(None):
        ctypes.windll.kernel32.GlobalFree(handle)
        raise OSError("OpenClipboard failed")

    try:
        ctypes.windll.user32.EmptyClipboard()
        if not ctypes.windll.user32.SetClipboardData(CF_UNICODETEXT, handle):
            raise OSError("SetClipboardData failed")
        handle = None
    finally:
        ctypes.windll.user32.CloseClipboard()
        if handle:
            ctypes.windll.kernel32.GlobalFree(handle)


def paste_text(text: str) -> None:
    set_clipboard_text(text)
    press_ctrl_shortcut(0x56)
    wait(0.2)


def clear_and_type(edit: auto.EditControl, text: str) -> None:
    edit.Click()
    clear_active_input()
    auto.SendKeys(text, waitTime=0.05)


def click_relative(win: auto.WindowControl, x_ratio: float, y_ratio: float, right: bool = False) -> tuple[int, int]:
    rect = safe_window_rect(win, win.Name)
    width = rect.right - rect.left
    height = rect.bottom - rect.top
    x = rect.left + int(width * x_ratio)
    y = rect.top + int(height * y_ratio)
    click_at(x, y, right=right)
    return x, y


def click_relative_offset(win: auto.WindowControl, x: int, y: int, right: bool = False) -> tuple[int, int]:
    rect = safe_window_rect(win, win.Name)
    click_at(rect.left + x, rect.top + y, right=right)
    return rect.left + x, rect.top + y


def search_chat_by_coordinates(win: auto.WindowControl, chat_name: str, x_ratio: float, y_ratio: float, settle_delay: float) -> None:
    click_relative(win, x_ratio, y_ratio, right=False)
    wait(0.4)
    clear_active_input()
    paste_text(chat_name)
    wait(settle_delay)
    # 新版微信会先把焦点落到“网络/AI 搜索”，第二次向下才是群聊结果。
    auto.SendKeys("{Down}{Down}{Enter}", waitTime=0.05)
    wait(settle_delay)


def capture_window_image(win: auto.WindowControl) -> tuple[ImageGrab.Image, auto.Rect]:
    rect = safe_window_rect(win, win.Name)
    bbox = (rect.left, rect.top, rect.right, rect.bottom)
    return ImageGrab.grab(bbox=bbox, all_screens=True), rect


def load_state(path: Path) -> dict:
    if not path.exists():
        return {"sent_pairs": {}}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"sent_pairs": {}}


def save_state(path: Path, state: dict) -> None:
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def state_key(source_chat: str) -> str:
    return f"{source_chat}:{date.today().isoformat()}"


def seen_signatures_for_today(state: dict, source_chat: str) -> set[str]:
    return set(state.get("sent_pairs", {}).get(state_key(source_chat), []))


def mark_signature_sent(state: dict, source_chat: str, signature: str) -> None:
    key = state_key(source_chat)
    sent_pairs = state.setdefault("sent_pairs", {})
    values = sent_pairs.setdefault(key, [])
    if signature not in values:
        values.append(signature)


def crop_chat_area(image) -> tuple[np.ndarray, tuple[int, int]]:
    arr = np.array(image)
    height, width = arr.shape[:2]
    x1 = int(width * 0.40)
    y1 = int(height * 0.08)
    return arr[y1:height, x1:width], (x1, y1)


def detect_bottom_video_center(image) -> tuple[int, int, int, int] | None:
    chat_arr, (offset_x, offset_y) = crop_chat_area(image)
    gray = cv2.cvtColor(chat_arr, cv2.COLOR_RGB2GRAY)
    circles = cv2.HoughCircles(
        gray,
        cv2.HOUGH_GRADIENT,
        dp=1.2,
        minDist=80,
        param1=100,
        param2=25,
        minRadius=18,
        maxRadius=45,
    )
    if circles is None:
        return None

    candidates: list[tuple[int, int, int]] = []
    for x, y, r in circles[0]:
        if x < chat_arr.shape[1] * 0.45:
            continue
        # 只看聊天区中部靠右的位置，过滤头像和底部工具栏里的圆形元素。
        if y < chat_arr.shape[0] * 0.15 or y > chat_arr.shape[0] * 0.68:
            continue
        if x > chat_arr.shape[1] * 0.90:
            continue
        candidates.append((int(x), int(y), int(r)))

    if not candidates:
        return None

    x, y, r = max(candidates, key=lambda item: item[1])
    return offset_x + x, offset_y + y, r, offset_y


def find_text_bubble_info(image, video_center: tuple[int, int, int, int]) -> tuple[tuple[int, int, int, int] | None, str]:
    arr = np.array(image)
    height, width = arr.shape[:2]
    cx, cy, _, chat_top = video_center

    result, _ = OCR_ENGINE(arr)
    if result:
        text_boxes: list[tuple[int, int, int, int]] = []
        text_map: list[tuple[tuple[int, int, int, int], str]] = []
        for item in result:
            points = item[0]
            text = item[1].strip()
            if not text or len(text) < 4:
                continue
            if re.fullmatch(r"[\d:]+", text):
                continue
            xs = [int(p[0]) for p in points]
            ys = [int(p[1]) for p in points]
            left, top, right, bottom = min(xs), min(ys), max(xs), max(ys)
            if right < int(width * 0.52):
                continue
            if top < chat_top + int(height * 0.05):
                continue
            if bottom >= cy or bottom < cy - int(height * 0.60):
                continue
            text_boxes.append((left, top, right, bottom))
            text_map.append(((left, top, right, bottom), text))

        if text_boxes:
            anchor = max(text_boxes, key=lambda box: box[3])
            cluster = [
                box
                for box in text_boxes
                if abs(box[3] - anchor[3]) <= 90 or abs(box[1] - anchor[1]) <= 90
            ]
            cluster_texts = [text for box, text in text_map if box in cluster]
            left = min(box[0] for box in cluster) - 12
            top = min(box[1] for box in cluster) - 12
            right = max(box[2] for box in cluster) + 12
            bottom = max(box[3] for box in cluster) + 12
            return (
                max(0, left),
                max(chat_top, top),
                min(width, right),
                min(height, bottom),
            ), " ".join(cluster_texts)

    search_top = max(chat_top, cy - int(height * 0.28))
    search_bottom = max(search_top + 20, cy - int(height * 0.08))
    search_left = int(width * 0.45)
    search_right = int(width * 0.95)
    region = arr[search_top:search_bottom, search_left:search_right]
    if region.size == 0:
        return None

    mask = (
        (region[:, :, 0] >= 130)
        & (region[:, :, 0] <= 210)
        & (region[:, :, 1] >= 210)
        & (region[:, :, 2] >= 120)
        & (region[:, :, 2] <= 210)
    )
    mask = (mask.astype(np.uint8) * 255)
    # 文字会把绿色气泡打断，闭运算可以把这些空洞补起来。
    kernel = np.ones((9, 9), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = [c for c in contours if cv2.contourArea(c) > 1000]
    if not contours:
        return None, ""

    contour = max(contours, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(contour)
    return (search_left + x, search_top + y, search_left + x + w, search_top + y + h), ""


def normalize_pair_text(text: str) -> str:
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^\w\u4e00-\u9fff]", "", text)
    return text[:80]


def normalize_ocr_text(text: str) -> str:
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^\w\u4e00-\u9fff()（）]", "", text)
    return text


def build_pair_signature(
    image,
    text_bbox: tuple[int, int, int, int],
    video_center: tuple[int, int, int, int],
    text_content: str,
) -> str:
    normalized_text = normalize_pair_text(text_content)
    if normalized_text:
        return normalized_text

    cx, cy, r, _ = video_center
    video_bbox = (
        max(0, cx - int(r * 2.8)),
        max(0, cy - int(r * 3.8)),
        min(image.width, cx + int(r * 2.8)),
        min(image.height, cy + int(r * 4.2)),
    )
    text_crop = image.crop(text_bbox)
    video_crop = image.crop(video_bbox)
    return f"{imagehash.phash(text_crop)}:{imagehash.phash(video_crop)}"


def detect_visible_pair_from_image(image, rect: auto.Rect) -> PairCandidate | None:
    video_center = detect_bottom_video_center(image)
    if video_center is None:
        return None

    text_bbox, text_content = find_text_bubble_info(image, video_center)
    if text_bbox is None:
        return None

    signature = build_pair_signature(image, text_bbox, video_center, text_content)
    tx1, ty1, tx2, ty2 = text_bbox
    text_x = (tx1 + tx2) // 2
    text_y = (ty1 + ty2) // 2
    vx, vy, _, _ = video_center

    width = rect.right - rect.left
    height = rect.bottom - rect.top
    menu_offset = max(26, min(320, vy - text_y + 62))

    return PairCandidate(
        text_x_ratio=text_x / width,
        text_y_ratio=text_y / height,
        video_x_ratio=vx / width,
        video_y_ratio=vy / height,
        text_menu_offset=menu_offset,
        signature=signature,
        text_content=text_content,
        approx_top_ratio=ty1 / height,
    )


def detect_visible_pair(win: auto.WindowControl) -> PairCandidate | None:
    image, rect = capture_window_image(win)
    return detect_visible_pair_from_image(image, rect)


def contains_old_day_marker_in_image(image) -> bool:
    result, _ = OCR_ENGINE(np.array(image))
    if not result:
        return False

    today = date.today()
    old_day_patterns = [
        "昨天",
        "星期",
        "周",
        f"{today.year - 1}年",
        f"{today.month - 1}月" if today.month > 1 else "",
    ]
    current_date_pattern = re.compile(rf"{today.year}年0?{today.month}月0?{today.day}日")

    for item in result:
        text = item[1]
        if current_date_pattern.search(text):
            continue
        if any(token and token in text for token in old_day_patterns):
            return True
        if re.search(r"\d{4}年\d{1,2}月\d{1,2}日", text) and not current_date_pattern.search(text):
            return True
    return False


def contains_send_dialog_in_image(image) -> bool:
    result, _ = OCR_ENGINE(np.array(image))
    if not result:
        return False
    return any("发送给" in item[1] for item in result)


def find_ocr_text_center(image, keyword: str) -> tuple[int, int] | None:
    result, _ = OCR_ENGINE(np.array(image))
    if not result:
        return None
    for item in result:
        text = item[1]
        if keyword not in text:
            continue
        points = item[0]
        xs = [int(p[0]) for p in points]
        ys = [int(p[1]) for p in points]
        return (min(xs) + max(xs)) // 2, (min(ys) + max(ys)) // 2
    return None


def find_chat_result_center(image, chat_name: str) -> tuple[int, int] | None:
    result, _ = OCR_ENGINE(np.array(image))
    if not result:
        return None

    target = normalize_ocr_text(chat_name)
    best: tuple[tuple[int, int], int] | None = None
    height, width = np.array(image).shape[:2]

    for item in result:
        text = normalize_ocr_text(item[1])
        if not text:
            continue
        if target not in text and text not in target:
            continue
        points = item[0]
        xs = [int(p[0]) for p in points]
        ys = [int(p[1]) for p in points]
        left, top, right, bottom = min(xs), min(ys), max(xs), max(ys)
        if left > width * 0.55:
            continue
        if top < height * 0.12 or bottom > height * 0.90:
            continue
        center = ((left + right) // 2, (top + bottom) // 2)
        score = len(set(text) & set(target))
        if best is None or score > best[1]:
            best = (center, score)

    return best[0] if best else None


def find_sidebar_chat_center(image, chat_name: str) -> tuple[int, int] | None:
    result, _ = OCR_ENGINE(np.array(image))
    if not result:
        return None

    target = normalize_ocr_text(chat_name)
    best: tuple[tuple[int, int], int] | None = None
    height, width = np.array(image).shape[:2]

    for item in result:
        text = normalize_ocr_text(item[1])
        if not text:
            continue
        if target not in text and text not in target:
            continue
        points = item[0]
        xs = [int(p[0]) for p in points]
        ys = [int(p[1]) for p in points]
        left, top, right, bottom = min(xs), min(ys), max(xs), max(ys)
        if right > width * 0.38:
            continue
        if top < height * 0.12 or bottom > height * 0.92:
            continue
        center = ((left + right) // 2, (top + bottom) // 2)
        score = len(set(text) & set(target))
        if best is None or score > best[1]:
            best = (center, score)

    return best[0] if best else None


def contains_old_day_marker(win: auto.WindowControl) -> bool:
    image, _ = capture_window_image(win)
    return contains_old_day_marker_in_image(image)


def scroll_history_up(win: auto.WindowControl, delta: int, settle_delay: float) -> None:
    # 把焦点放到右侧滚动条附近，避免滚轮落在视频卡片上触发预览/缩放。
    click_relative(win, 0.985, 0.50)
    wait(0.2)
    mouse_wheel(delta)
    wait(settle_delay)


def open_chat(
    win: auto.WindowControl,
    chat_name: str,
    search_box_x_ratio: float,
    search_box_y_ratio: float,
    settle_delay: float,
) -> None:
    activate_wechat(win)

    image, rect = capture_window_image(win)
    sidebar_center = find_sidebar_chat_center(image, chat_name)
    if sidebar_center is not None:
        width = rect.right - rect.left
        click_at(rect.left + int(width * 0.18), rect.top + sidebar_center[1], right=False)
        wait(settle_delay)
        return

    try:
        search_chat_by_coordinates(
            win,
            chat_name,
            search_box_x_ratio,
            search_box_y_ratio,
            settle_delay,
        )
        return
    except Exception:
        logging.info("坐标搜索失败，回退到 UI 自动化搜索。")

    control = first_existing(
        [
            win.EditControl(Name="搜索"),
            win.EditControl(AutomationId="SearchBox"),
            win.EditControl(searchDepth=6),
        ]
    )
    if control is None:
        raise RuntimeError("未找到微信搜索入口，请确认微信主窗口处于前台。")

    search_box = control  # type: ignore[assignment]
    clear_and_type(search_box, chat_name)
    wait(settle_delay)

    candidate = first_existing(
        [
            win.ListItemControl(Name=chat_name, searchDepth=8),
            win.TextControl(Name=chat_name, searchDepth=8),
        ]
    )
    if candidate is None:
        raise RuntimeError(f"未找到聊天: {chat_name}")

    candidate.Click(simulateMove=False)
    wait(settle_delay)


def message_candidates(win: auto.WindowControl, max_scan: int) -> list[auto.Control]:
    controls = list(win.GetChildren())
    descendants: list[auto.Control] = []
    queue = controls[:]
    while queue and len(descendants) < max_scan * 8:
        current = queue.pop(0)
        descendants.append(current)
        try:
            queue.extend(current.GetChildren())
        except Exception:
            continue

    candidates: list[auto.Control] = []
    for control in descendants:
        try:
            rect = control.BoundingRectangle
            if rect.width() < 40 or rect.height() < 20:
                continue
            name = (control.Name or "").strip()
            class_name = (control.ClassName or "").strip()
            if any(token in class_name for token in ("Chat", "Message", "Msg")):
                candidates.append(control)
                continue
            if name and len(name) >= 2:
                candidates.append(control)
        except Exception:
            continue

    candidates.sort(key=lambda ctrl: ctrl.BoundingRectangle.bottom, reverse=True)
    return candidates[:max_scan]


def pick_latest_message(win: auto.WindowControl, max_scan: int) -> auto.Control:
    candidates = message_candidates(win, max_scan)
    if not candidates:
        raise RuntimeError("没有找到可转发的消息控件。")
    return candidates[0]


def focus_latest_visible_message(win: auto.WindowControl, x_ratio: float, y_ratio: float, delay: float) -> None:
    click_relative(win, x_ratio, y_ratio, right=False)
    wait(delay)


def click_menu_item(root: auto.Control, names: list[str], timeout: float = 5) -> auto.Control:
    deadline = time.time() + timeout
    while time.time() < deadline:
        for name in names:
            item = first_existing(
                [
                    root.MenuItemControl(Name=name, searchDepth=8),
                    root.ButtonControl(Name=name, searchDepth=8),
                    root.TextControl(Name=name, searchDepth=8),
                ]
            )
            if item is not None:
                item.Click(simulateMove=False)
                return item
        wait(0.2)
    raise RuntimeError(f"未找到菜单项: {'/'.join(names)}")


def open_forward_dialog(win: auto.WindowControl, message: auto.Control, delay: float) -> None:
    message.RightClick(simulateMove=False)
    wait(delay)
    try:
        click_menu_item(win, ["转发...", "转发"])
    except Exception:
        # 这版微信的右键菜单不一定暴露给 UI 自动化，首项就是“转发”，直接回车更稳。
        auto.SendKeys("{Enter}", waitTime=0.05)
    wait(delay)


def click_context_menu_action_by_coordinates(
    win: auto.WindowControl,
    x_ratio: float,
    y_ratio: float,
    menu_y_offset: int,
    delay: float,
) -> None:
    rect = safe_window_rect(win, win.Name)
    bbox = (rect.left, rect.top, rect.right, rect.bottom)
    before = ImageGrab.grab(bbox=bbox, all_screens=True)
    x, y = click_relative(win, x_ratio, y_ratio, right=True)
    wait(delay)
    after = ImageGrab.grab(bbox=bbox, all_screens=True)
    diff = ImageChops.difference(before, after).convert("L")
    mask = diff.point(lambda p: 255 if p > 25 else 0)
    diff_rect = mask.getbbox()
    if diff_rect:
        left, top, right, _ = diff_rect
        click_at(rect.left + (left + right) // 2, rect.top + top + menu_y_offset, right=False)
    else:
        auto.SendKeys("{Enter}", waitTime=0.05)
    wait(delay)


def open_context_menu_and_get_bounds(
    win: auto.WindowControl,
    x_ratio: float,
    y_ratio: float,
    delay: float,
) -> tuple[auto.Rect, tuple[int, int, int, int] | None]:
    rect = safe_window_rect(win, win.Name)
    bbox = (rect.left, rect.top, rect.right, rect.bottom)
    before = ImageGrab.grab(bbox=bbox, all_screens=True)
    click_relative(win, x_ratio, y_ratio, right=True)
    wait(delay)
    after = ImageGrab.grab(bbox=bbox, all_screens=True)
    diff = ImageChops.difference(before, after).convert("L")
    mask = diff.point(lambda p: 255 if p > 25 else 0)
    return rect, mask.getbbox()


def click_context_menu_row(
    win: auto.WindowControl,
    menu_bounds: tuple[int, int, int, int] | None,
    row_index: int,
    delay: float,
) -> None:
    if not menu_bounds:
        auto.SendKeys("{Esc}", waitTime=0.05)
        wait(0.2)
        return
    rect = safe_window_rect(win, win.Name)
    left, top, right, bottom = menu_bounds
    row_count = 10
    row_height = max(24, (bottom - top) // row_count)
    click_x = rect.left + (left + right) // 2
    click_y = rect.top + top + row_height // 2 + row_height * row_index
    click_at(click_x, click_y, right=False)
    wait(delay)


def click_context_menu_ratio(
    win: auto.WindowControl,
    menu_bounds: tuple[int, int, int, int] | None,
    y_ratio: float,
    delay: float,
) -> None:
    if not menu_bounds:
        auto.SendKeys("{Esc}", waitTime=0.05)
        wait(0.2)
        return
    rect = safe_window_rect(win, win.Name)
    left, top, right, bottom = menu_bounds
    click_x = rect.left + (left + right) // 2
    click_y = rect.top + top + int((bottom - top) * y_ratio)
    click_at(click_x, click_y, right=False)
    wait(delay)


def open_forward_dialog_by_coordinates(
    win: auto.WindowControl,
    x_ratio: float,
    y_ratio: float,
    menu_y_offset: int,
    delay: float,
) -> None:
    click_context_menu_action_by_coordinates(win, x_ratio, y_ratio, menu_y_offset, delay)


def forward_single_message(config: Config, x_ratio: float, y_ratio: float, menu_y_offset: int) -> None:
    win = window(config.wechat_window_title)
    activate_wechat(win)
    focus_latest_visible_message(win, x_ratio, y_ratio, config.settle_delay_seconds)
    open_forward_dialog_by_coordinates(
        win,
        x_ratio,
        y_ratio,
        menu_y_offset,
        config.dialog_delay_seconds,
    )

    if config.dry_run or config.preview_forward_dialog:
        logging.info("安全模式：已打开转发流程，但不会执行最终发送。")
        return

    forward_to_targets(config)
    logging.info("单条消息转发完成。")


def open_multi_select_for_pair(config: Config) -> None:
    win = window(config.wechat_window_title)
    activate_wechat(win)
    focus_latest_visible_message(
        win,
        config.text_message_click_x_ratio,
        config.text_message_click_y_ratio,
        config.settle_delay_seconds,
    )
    _, menu_bounds = open_context_menu_and_get_bounds(
        win,
        config.text_message_click_x_ratio,
        config.text_message_click_y_ratio,
        config.dialog_delay_seconds,
    )
    click_context_menu_ratio(
        win,
        menu_bounds,
        config.multi_select_menu_click_ratio,
        config.dialog_delay_seconds,
    )


def open_multi_select_for_candidate(config: Config, pair: PairCandidate) -> None:
    win = ensure_wechat_window(config.wechat_window_title)
    activate_wechat(win)
    focus_latest_visible_message(
        win,
        pair.text_x_ratio,
        pair.text_y_ratio,
        config.settle_delay_seconds,
    )
    _, menu_bounds = open_context_menu_and_get_bounds(
        win,
        pair.text_x_ratio,
        pair.text_y_ratio,
        config.dialog_delay_seconds,
    )
    click_context_menu_ratio(
        win,
        menu_bounds,
        config.multi_select_menu_click_ratio,
        config.dialog_delay_seconds,
    )


def resolve_single_pair_candidate(config: Config, win: auto.WindowControl) -> PairCandidate:
    for _ in range(5):
        pair = detect_visible_pair(win)
        if pair is not None:
            return pair
        wait(0.5)

    logging.info("未动态识别到文字+视频组合，回退到配置坐标，但仍只点击左侧勾选圈。")
    return PairCandidate(
        text_x_ratio=config.text_message_click_x_ratio,
        text_y_ratio=config.text_message_click_y_ratio,
        video_x_ratio=config.multi_select_checkbox_x_ratio,
        video_y_ratio=config.grouped_video_click_y_ratio,
        text_menu_offset=config.multi_select_menu_y_offset,
        signature="single-run-fallback",
        text_content="",
        approx_top_ratio=min(config.text_message_click_y_ratio, config.grouped_video_click_y_ratio),
    )


def select_grouped_video_message(config: Config) -> None:
    win = window(config.wechat_window_title)
    activate_wechat(win)
    click_relative(
        win,
        config.multi_select_checkbox_x_ratio,
        config.grouped_video_click_y_ratio,
        right=False,
    )
    wait(config.dialog_delay_seconds)


def select_grouped_video_for_candidate(config: Config, pair: PairCandidate) -> None:
    win = ensure_wechat_window(config.wechat_window_title)
    activate_wechat(win)
    # 多选模式下点击左侧勾选圈，避免点到视频内容本体触发预览。
    click_relative(
        win,
        config.multi_select_checkbox_x_ratio,
        pair.video_y_ratio,
        right=False,
    )
    wait(config.dialog_delay_seconds)


def open_forward_dialog_from_multi_select(config: Config) -> None:
    win = ensure_wechat_window(config.wechat_window_title)
    activate_wechat(win)
    click_relative(
        win,
        config.multi_forward_button_x_ratio,
        config.multi_forward_button_y_ratio,
        right=False,
    )
    wait(config.dialog_delay_seconds)


def close_forward_dialog_without_sending(config: Config) -> None:
    win = ensure_wechat_window(config.wechat_window_title)
    activate_wechat(win)
    image, rect = capture_window_image(win)
    cancel_center = find_ocr_text_center(image, "取消")
    if cancel_center is not None:
        click_at(rect.left + cancel_center[0], rect.top + cancel_center[1], right=False)
    else:
        click_relative(
            win,
            config.forward_dialog_cancel_x_ratio,
            config.forward_dialog_cancel_y_ratio,
            right=False,
        )
    wait(0.5)
    image, _ = capture_window_image(win)
    if contains_send_dialog_in_image(image):
        press_key(0x1B)
    wait(config.dialog_delay_seconds)


def cleanup_forward_dialog_if_present(config: Config) -> None:
    win = ensure_wechat_window(config.wechat_window_title)
    activate_wechat(win)
    image, _ = capture_window_image(win)
    if contains_send_dialog_in_image(image):
        close_forward_dialog_without_sending(config)


def reopen_source_chat(config: Config) -> auto.WindowControl:
    deadline = time.time() + 6
    last_error: Exception | None = None
    while time.time() < deadline:
        try:
            win = ensure_wechat_window(config.wechat_window_title, timeout_seconds=2)
            activate_wechat(win)
            open_chat(
                win,
                config.source_chat,
                config.search_box_x_ratio,
                config.search_box_y_ratio,
                config.settle_delay_seconds,
            )
            clear_chat_input_if_needed(win)
            return win
        except Exception as exc:
            last_error = exc
            wait(0.5)
    raise RuntimeError(str(last_error) if last_error else "未能重新打开来源群")


def clear_chat_input_if_needed(win: auto.WindowControl) -> None:
    rect = safe_window_rect(win, win.Name)
    width = rect.right - rect.left
    height = rect.bottom - rect.top
    # 点击输入框，再清空可能误留的搜索/测试字符。
    click_at(rect.left + int(width * 0.55), rect.top + int(height * 0.84), right=False)
    wait(0.2)
    clear_active_input()
    wait(0.2)


def clear_forward_dialog_search(win: auto.WindowControl, clear_x_ratio: float, search_y_ratio: float, delay: float) -> None:
    click_relative(win, clear_x_ratio, search_y_ratio, right=False)
    wait(0.2)
    press_key(0x08)
    wait(delay)


def forward_to_targets(config: Config) -> None:
    win = ensure_wechat_window(config.wechat_window_title)
    activate_wechat(win)

    for chat_name in config.target_chats:
        click_relative(win, config.forward_dialog_search_x_ratio, config.forward_dialog_search_y_ratio)
        wait(0.2)
        clear_forward_dialog_search(
            win,
            config.forward_dialog_clear_x_ratio,
            config.forward_dialog_search_y_ratio,
            config.dialog_delay_seconds,
        )
        click_relative(win, config.forward_dialog_search_x_ratio, config.forward_dialog_search_y_ratio)
        wait(0.2)
        paste_text(chat_name)
        wait(config.dialog_delay_seconds)
        image, rect = capture_window_image(win)
        target_center = find_chat_result_center(image, chat_name)
        if target_center is not None:
            click_at(rect.left + target_center[0], rect.top + target_center[1], right=False)
        else:
            click_relative(win, config.forward_dialog_result_x_ratio, config.forward_dialog_result_y_ratio)
        wait(0.5)

    image, rect = capture_window_image(win)
    send_center = find_ocr_text_center(image, "发送")
    if send_center is not None:
        click_at(rect.left + send_center[0], rect.top + send_center[1], right=False)
    else:
        click_relative(win, config.forward_dialog_send_x_ratio, config.forward_dialog_send_y_ratio)
    wait(config.dialog_delay_seconds)


def dump_controls(win: auto.WindowControl, path: Path) -> None:
    lines: list[str] = []
    queue: list[tuple[int, auto.Control]] = [(0, win)]
    while queue and len(lines) < 400:
        depth, current = queue.pop(0)
        try:
            rect = current.BoundingRectangle
            lines.append(
                f"{'  ' * depth}{current.ControlTypeName} | Name={current.Name!r} | "
                f"Class={current.ClassName!r} | Rect=({rect.left},{rect.top},{rect.right},{rect.bottom})"
            )
            for child in current.GetChildren():
                queue.append((depth + 1, child))
        except Exception:
            continue
    path.write_text("\n".join(lines), encoding="utf-8")


def run(config: Config) -> None:
    win = window(config.wechat_window_title)
    activate_wechat(win)
    cleanup_forward_dialog_if_present(config)
    win = ensure_wechat_window(config.wechat_window_title)
    activate_wechat(win)
    open_chat(
        win,
        config.source_chat,
        config.search_box_x_ratio,
        config.search_box_y_ratio,
        config.settle_delay_seconds,
    )

    if config.batch_mode_enabled:
        state_path = Path(config.state_file)
        state = load_state(state_path)
        seen_today = seen_signatures_for_today(state, config.source_chat)
        seen_this_run: set[str] = set()
        processed = 0
        seen_streak = 0
        empty_streak = 0
        scrolls = 0

        while processed < config.max_pairs_per_run and scrolls < config.max_scroll_pages:
            pair = detect_visible_pair(win)
            if pair is None:
                empty_streak += 1
                logging.info("当前页没有识别到新的文字+视频组合，继续上翻。")
                if empty_streak >= config.stop_after_empty_pages:
                    break
                scroll_history_up(win, config.scroll_wheel_delta, config.settle_delay_seconds)
                scrolls += 1
                continue

            empty_streak = 0
            if pair.signature in seen_today or pair.signature in seen_this_run:
                seen_streak += 1
                logging.info("识别到已发送组合，连续命中 %s 次。", seen_streak)
                if seen_streak >= config.stop_after_seen_pairs:
                    break
            else:
                seen_streak = 0
                open_multi_select_for_candidate(config, pair)
                select_grouped_video_for_candidate(config, pair)
                open_forward_dialog_from_multi_select(config)
                if config.dry_run or config.preview_forward_dialog:
                    logging.info("安全模式：已打开组合转发流程，关闭后继续扫描。")
                    seen_this_run.add(pair.signature)
                    close_forward_dialog_without_sending(config)
                    win = reopen_source_chat(config)
                else:
                    forward_to_targets(config)
                    mark_signature_sent(state, config.source_chat, pair.signature)
                    save_state(state_path, state)
                    seen_today = seen_signatures_for_today(state, config.source_chat)
                    seen_this_run.add(pair.signature)
                    processed += 1
                    logging.info("批量模式已转发 %s 组。", processed)
                    win = reopen_source_chat(config)

            if contains_old_day_marker(win):
                logging.info("已检测到旧日期标记，停止继续上翻。")
                break

            scroll_history_up(win, config.scroll_wheel_delta, config.settle_delay_seconds)
            scrolls += 1

        logging.info("批量扫描完成，本次共处理 %s 组。", processed)
        return

    if config.forward_as_grouped_pair:
        pair = resolve_single_pair_candidate(config, win)
        open_multi_select_for_candidate(config, pair)
        select_grouped_video_for_candidate(config, pair)
        open_forward_dialog_from_multi_select(config)
        if config.dry_run or config.preview_forward_dialog:
            logging.info("瀹夊叏妯″紡锛氬凡鎵撳紑缁勫悎杞彂娴佺▼锛屼絾涓嶄細鎵ц鏈€缁堝彂閫併€?")
            return
        forward_to_targets(config)
        logging.info("鏂囧瓧+瑙嗛缁勫悎宸查€愭潯杞彂瀹屾垚銆?")
        return

    message = None
    try:
        message = pick_latest_message(win, config.max_message_scan)
    except Exception:
        logging.info("未识别到消息控件，回退到坐标方式选择最新可见消息。")
        focus_latest_visible_message(
            win,
            config.message_click_x_ratio,
            config.message_click_y_ratio,
            config.settle_delay_seconds,
        )

    if config.forward_as_grouped_pair:
        open_multi_select_for_pair(config)
        select_grouped_video_message(config)
        open_forward_dialog_from_multi_select(config)
        if config.dry_run or config.preview_forward_dialog:
            logging.info("安全模式：已打开组合转发流程，但不会执行最终发送。")
            return
        forward_to_targets(config)
        logging.info("文字+视频组合已逐条转发完成。")
        return

    if message is not None and not config.forward_text_then_video:
        open_forward_dialog(win, message, config.dialog_delay_seconds)
        if config.dry_run or config.preview_forward_dialog:
            logging.info("安全模式：已打开转发流程，但不会执行最终发送。")
            return
        forward_to_targets(config)
        logging.info("转发完成。")
        return

    if config.forward_text_then_video:
        forward_single_message(
            config,
            config.text_message_click_x_ratio,
            config.text_message_click_y_ratio,
            config.text_forward_menu_y_offset,
        )
        wait(config.dialog_delay_seconds)
        open_chat(
            win,
            config.source_chat,
            config.search_box_x_ratio,
            config.search_box_y_ratio,
            config.settle_delay_seconds,
        )
        forward_single_message(
            config,
            config.message_click_x_ratio,
            config.message_click_y_ratio,
            config.video_forward_menu_y_offset,
        )
        logging.info("文字和视频已逐条转发完成。")
        return

    open_forward_dialog_by_coordinates(
        win,
        config.message_click_x_ratio,
        config.message_click_y_ratio,
        config.video_forward_menu_y_offset,
        config.dialog_delay_seconds,
    )
    if config.dry_run or config.preview_forward_dialog:
        logging.info("安全模式：已打开转发流程，但不会执行最终发送。")
        return
    forward_to_targets(config)
    logging.info("转发完成。")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="微信 PC 版群消息自动转发")
    parser.add_argument(
        "--config",
        default="config.json",
        help="配置文件路径，默认读取当前目录下的 config.json",
    )
    parser.add_argument(
        "--dump-ui",
        action="store_true",
        help="导出微信界面控件树，便于调试定位失败问题",
    )
    return parser.parse_args()


def main() -> int:
    setup_logging()
    args = parse_args()
    config_path = Path(args.config)

    if not config_path.exists():
        logging.error("配置文件不存在: %s", config_path)
        logging.info("请先复制 config.example.json 为 config.json，并填写群名。")
        return 1

    config = load_config(config_path)

    try:
        win = ensure_wechat_window(config.wechat_window_title)
        if args.dump_ui:
            dump_controls(win, Path("wechat_ui_dump.txt"))
            logging.info("已导出控件树到 wechat_ui_dump.txt")
            return 0

        run(config)
        return 0
    except Exception as exc:
        logging.exception("执行失败: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
