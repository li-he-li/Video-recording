"""
Automatic course recorder.

Features:
1. Full-screen recording with ffmpeg.
2. Detect lesson completion by comparing a screen region to a reference image.
3. Click "Next lesson" automatically.
4. Click "Play" automatically for the next lesson.
"""

import json
import os
import re
import shutil
import subprocess
import time
from datetime import datetime

import cv2
import numpy as np
import pyautogui


DEFAULT_AUDIO_DEVICE = ""
AUDIO_SYNC_FILTER = "aresample=async=1:min_hard_comp=0.100:first_pts=0"
DEFAULT_RECORDING_MODE = "ffmpeg"
SUPPORTED_RECORDING_MODES = {"ffmpeg", "nvidia"}


class Config:
    """Config file helper."""

    CONFIG_FILE = "recorder_config.json"

    @staticmethod
    def load():
        if os.path.exists(Config.CONFIG_FILE):
            with open(Config.CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    @staticmethod
    def save(config):
        with open(Config.CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)


class CourseRecorder:
    def __init__(self):
        self.config = None
        self.current_file_number = 1
        self.recording_process = None
        self.is_recording = False
        self.playing_state = 1  # 1 = playing, 0 = waiting to start
        self.current_output_file = None
        self.useless_skip_enabled = False
        self.useless_skip_runtime_config = None
        self.useless_skip_templates = []

    @staticmethod
    def _wait_for_yes(prompt):
        while True:
            print(prompt, end="", flush=True)
            answer = input().strip().lower()
            if answer == "y":
                return True
            print("Please input 'y' to confirm.")

    @staticmethod
    def _ask_yes_no(prompt, default=None):
        """Prompt y/n and return bool."""
        normalized_default = None
        if isinstance(default, str) and default.lower() in {"y", "n"}:
            normalized_default = default.lower()

        while True:
            print(prompt, end="", flush=True)
            answer = input().strip().lower()
            if not answer and normalized_default:
                answer = normalized_default
            if answer in {"y", "n"}:
                return answer == "y"
            print("Please input 'y' or 'n'.")

    @staticmethod
    def _input_positive_int(prompt, default):
        while True:
            print(prompt, end="", flush=True)
            answer = input().strip()
            if not answer:
                return default
            if answer.isdigit() and int(answer) > 0:
                return int(answer)
            print("Please input a positive integer.")

    @staticmethod
    def _input_float_in_range(prompt, default, min_value, max_value):
        while True:
            print(prompt, end="", flush=True)
            answer = input().strip()
            if not answer:
                return default
            try:
                value = float(answer)
            except ValueError:
                print("Please input a valid number.")
                continue
            if min_value <= value <= max_value:
                return value
            print(f"Please input a number in [{min_value}, {max_value}].")

    @staticmethod
    def _is_image_file(path):
        ext = os.path.splitext(path)[1].lower()
        return ext in {".png", ".jpg", ".jpeg", ".bmp", ".webp"}

    @staticmethod
    def _normalize_audio_device(audio_device):
        if audio_device is None:
            return ""
        value = str(audio_device).strip()
        if not value:
            return ""
        if value.lower().startswith("audio="):
            return value
        return f"audio={value}"

    @staticmethod
    def _extract_audio_device_name(audio_device):
        value = CourseRecorder._normalize_audio_device(audio_device)
        if value.lower().startswith("audio="):
            return value[6:].strip()
        return value.strip()

    @staticmethod
    def _decode_ffmpeg_output(raw_bytes):
        for encoding in ("utf-8", "gbk", "cp936"):
            try:
                return raw_bytes.decode(encoding)
            except UnicodeDecodeError:
                continue
        return raw_bytes.decode("latin1", errors="ignore")

    @classmethod
    def _list_audio_devices(cls):
        cmd = ["ffmpeg", "-hide_banner", "-list_devices", "true", "-f", "dshow", "-i", "dummy"]
        try:
            result = subprocess.run(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                check=False,
            )
        except FileNotFoundError:
            return []

        output = cls._decode_ffmpeg_output(result.stderr)
        devices = []
        in_audio_section = False
        saw_audio_section = False

        for line in output.splitlines():
            if "DirectShow audio devices" in line:
                in_audio_section = True
                saw_audio_section = True
                continue
            if in_audio_section and "DirectShow video devices" in line:
                break

            if not in_audio_section:
                continue
            if "Alternative name" in line:
                continue

            match = re.search(r'"(.+?)"', line)
            if match:
                devices.append(match.group(1).strip())

        # Fallback: some ffmpeg builds/locales do not print the expected section markers.
        if not devices and not saw_audio_section:
            for line in output.splitlines():
                if "dshow" not in line.lower():
                    continue
                if "Alternative name" in line:
                    continue
                match = re.search(r'"(.+?)"', line)
                if match:
                    name = match.group(1).strip()
                    if name not in devices:
                        devices.append(name)

        return devices

    @staticmethod
    def _audio_device_score(name):
        lowered = name.lower()
        if (
            any(k in lowered for k in ("stereo mix", "stereomix", "what u hear", "loopback"))
            or "\u7acb\u4f53\u58f0\u6df7\u97f3" in name
        ):
            return 100
        if any(k in lowered for k in ("cable output", "vb-audio", "virtual cable")):
            return 90
        if any(k in lowered for k in ("speaker", "speakers", "output")) or "\u626c\u58f0\u5668" in name:
            return 80
        if (
            any(k in lowered for k in ("microphone", "mic", "array"))
            or "\u9ea6\u514b\u98ce" in name
            or "\u9635\u5217\u9ea6\u514b\u98ce" in name
        ):
            return 30
        return 0

    @classmethod
    def _select_preferred_audio_device(cls, device_names):
        if not device_names:
            return ""

        best_score = -1
        best_name = device_names[0]

        for name in device_names:
            score = cls._audio_device_score(name)
            if score > best_score:
                best_score = score
                best_name = name

        return best_name

    def _resolve_audio_device(self):
        configured = self._normalize_audio_device((self.config or {}).get("audio_device", ""))
        if configured:
            return configured

        available_devices = self._list_audio_devices()

        preferred_name = self._select_preferred_audio_device(available_devices)
        return self._normalize_audio_device(preferred_name)

    def _get_recording_mode(self):
        mode = str((self.config or {}).get("recording_mode", DEFAULT_RECORDING_MODE)).strip().lower()
        if mode in SUPPORTED_RECORDING_MODES:
            return mode
        return DEFAULT_RECORDING_MODE

    def _get_nvidia_hotkey(self):
        raw = (self.config or {}).get("nvidia_toggle_hotkey", ["alt", "f9"])
        if isinstance(raw, str):
            keys = [k.strip().lower() for k in raw.replace("+", " ").split() if k.strip()]
        elif isinstance(raw, list):
            keys = [str(k).strip().lower() for k in raw if str(k).strip()]
        else:
            keys = []

        if len(keys) < 2:
            return ["alt", "f9"]
        return keys

    def setup_config(self):
        """Interactive setup."""
        print("=" * 60)
        print("Course Recorder - Initial Setup")
        print("=" * 60)

        previous_config = self.config or {}
        config = {}

        print("\n[Step 1] Set the Next Lesson button position")
        print("Move your mouse to the Next Lesson button.")
        print("Then type 'y' in this terminal and press Enter.")
        print("(Press Ctrl+C to cancel)")
        self._wait_for_yes("\nInput 'y' to capture current position: ")
        x, y = pyautogui.position()
        config["next_button"] = {"x": x, "y": y}
        print(f"Captured: ({x}, {y})")

        print("\n[Step 2] Set the Play button position")
        print("Move your mouse to the Play button.")
        print("Then type 'y' in this terminal and press Enter.")
        print("(Press Ctrl+C to cancel)")
        self._wait_for_yes("\nInput 'y' to capture current position: ")
        x, y = pyautogui.position()
        config["play_button"] = {"x": x, "y": y}
        print(f"Captured: ({x}, {y})")

        print("\n[Step 3] Set the player status detection area")
        print("This area should cover a stable visual region that changes")
        print("when the lesson reaches the stopped/finished state.")
        print("(Press Ctrl+C to cancel)")

        print("\nMove mouse to TOP-LEFT corner of detection area.")
        self._wait_for_yes("Input 'y' to capture top-left: ")
        x1, y1 = pyautogui.position()
        print(f"Top-left: ({x1}, {y1})")

        print("\nMove mouse to BOTTOM-RIGHT corner of detection area.")
        self._wait_for_yes("Input 'y' to capture bottom-right: ")
        x2, y2 = pyautogui.position()
        print(f"Bottom-right: ({x2}, {y2})")

        config["detection_area"] = {
            "x1": min(x1, x2),
            "y1": min(y1, y2),
            "x2": max(x1, x2),
            "y2": max(y1, y2),
        }

        print("\n[Step 4] Capture reference image for FINISHED state")
        print("Make sure the player is in finished/stopped state, then capture.")
        self._wait_for_yes("Input 'y' to capture reference image: ")
        area = config["detection_area"]
        screenshot = pyautogui.screenshot(
            region=(area["x1"], area["y1"], area["x2"] - area["x1"], area["y2"] - area["y1"])
        )
        image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        cv2.imwrite("stopped_reference.png", image)
        config["reference_image"] = "stopped_reference.png"
        config["recording_mode"] = str(
            previous_config.get("recording_mode", DEFAULT_RECORDING_MODE)
        ).strip().lower()
        if config["recording_mode"] not in SUPPORTED_RECORDING_MODES:
            config["recording_mode"] = DEFAULT_RECORDING_MODE

        hotkey = previous_config.get("nvidia_toggle_hotkey", ["alt", "f9"])
        if isinstance(hotkey, list) and len(hotkey) >= 2:
            config["nvidia_toggle_hotkey"] = hotkey
        else:
            config["nvidia_toggle_hotkey"] = ["alt", "f9"]
        # Prefer system loopback devices (Stereo Mix/CABLE) to avoid muffled mic capture.
        config["audio_device"] = previous_config.get("audio_device") or self._resolve_audio_device() or DEFAULT_AUDIO_DEVICE
        if isinstance(previous_config.get("useless_page_skip"), dict):
            config["useless_page_skip"] = previous_config["useless_page_skip"]

        print("\nSetup completed.")
        print(f"Next button: {config['next_button']}")
        print(f"Play button: {config['play_button']}")
        print(f"Detection area: {config['detection_area']}")
        print(f"Reference image: {config['reference_image']}")
        print(f"Audio device: {config['audio_device']}")

        Config.save(config)
        self.config = config
        return True

    def setup_useless_page_skip_config(self):
        """Interactive setup for useless-page skip feature."""
        if self.config is None:
            self.config = {}

        previous = self.config.get("useless_page_skip", {})
        default_templates_dir = str(previous.get("templates_dir", "image")).strip() or "image"
        default_area_count = len(previous.get("areas", [])) or 1
        default_similarity = float(previous.get("similarity_threshold", 0.86))
        default_skip_next = previous.get("next_button")

        print("\n" + "=" * 60)
        print("Useless Page Skip - Setup")
        print("=" * 60)
        print("If any detection area looks similar to templates in folder 'image',")
        print("the script will click Next automatically.")

        print(f"\nTemplate folder path (default: {default_templates_dir}): ", end="", flush=True)
        templates_dir_input = input().strip()
        templates_dir = templates_dir_input or default_templates_dir
        if not os.path.isdir(templates_dir):
            os.makedirs(templates_dir, exist_ok=True)
            print(f"Created template folder: {templates_dir}")

        area_count = self._input_positive_int(
            f"How many detection areas? (default: {default_area_count}): ",
            default_area_count,
        )
        similarity_threshold = self._input_float_in_range(
            f"Similarity threshold [0.0-1.0] (default: {default_similarity:.2f}): ",
            default_similarity,
            0.0,
            1.0,
        )

        print("\nSet dedicated Next button position for useless-page skip.")
        if isinstance(default_skip_next, dict):
            print(
                f"Current useless-skip Next button: "
                f"({default_skip_next.get('x')}, {default_skip_next.get('y')})"
            )
        print("Move mouse to this button and type 'y'.")
        self._wait_for_yes("Input 'y' to capture useless-skip next-button position: ")
        skip_next_x, skip_next_y = pyautogui.position()
        print(f"Captured useless-skip Next button: ({skip_next_x}, {skip_next_y})")

        areas = []
        print("\nCapture each detection area by moving mouse and pressing 'y'.")
        for idx in range(area_count):
            print(f"\n[Area {idx + 1}/{area_count}] Move mouse to TOP-LEFT corner.")
            self._wait_for_yes("Input 'y' to capture top-left: ")
            x1, y1 = pyautogui.position()
            print(f"Top-left: ({x1}, {y1})")

            print("Move mouse to BOTTOM-RIGHT corner.")
            self._wait_for_yes("Input 'y' to capture bottom-right: ")
            x2, y2 = pyautogui.position()
            print(f"Bottom-right: ({x2}, {y2})")

            areas.append(
                {
                    "x1": min(x1, x2),
                    "y1": min(y1, y2),
                    "x2": max(x1, x2),
                    "y2": max(y1, y2),
                }
            )
            area = areas[-1]
            width = area["x2"] - area["x1"]
            height = area["y2"] - area["y1"]
            if width > 0 and height > 0:
                screenshot = pyautogui.screenshot(region=(area["x1"], area["y1"], width, height))
                image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                sample_name = f"area_{idx + 1}_capture_{stamp}.png"
                sample_path = os.path.join(templates_dir, sample_name)
                cv2.imwrite(sample_path, image)
                print(f"Saved area screenshot: {sample_path}")

        skip_config = {
            "enabled": True,
            "templates_dir": templates_dir,
            "areas": areas,
            "next_button": {"x": skip_next_x, "y": skip_next_y},
            "similarity_threshold": float(similarity_threshold),
            "check_interval_seconds": 13,
        }
        self.config["useless_page_skip"] = skip_config
        Config.save(self.config)

        print("\nUseless-page skip setup completed.")
        print(f"Template folder: {templates_dir}")
        print(f"Areas: {len(areas)}")
        print(f"Useless-skip Next button: ({skip_next_x}, {skip_next_y})")
        print(f"Similarity threshold: {skip_config['similarity_threshold']:.2f}")
        print("Check interval: 13s")
        return True

    def _load_useless_skip_templates(self, templates_dir):
        templates = []
        if not templates_dir or not os.path.isdir(templates_dir):
            return templates

        for name in sorted(os.listdir(templates_dir)):
            path = os.path.join(templates_dir, name)
            if not os.path.isfile(path) or not self._is_image_file(path):
                continue
            image = cv2.imread(path)
            if image is None:
                continue
            templates.append({"name": name, "path": path, "image": image})
        return templates

    def _get_useless_skip_check_interval(self):
        return 13

    def _prepare_useless_page_skip(self):
        """Ask if useless-page skip is enabled for this run and prepare runtime state."""
        self.useless_skip_enabled = False
        self.useless_skip_runtime_config = None
        self.useless_skip_templates = []

        existing = (self.config or {}).get("useless_page_skip", {})
        default_choice = "y" if bool(existing.get("enabled")) else "n"
        enable = self._ask_yes_no(
            f"Enable useless-page skip this run? (y/n, default {default_choice}): ",
            default=default_choice,
        )
        if not enable:
            print("Useless-page skip: OFF")
            return

        needs_setup = (
            not existing.get("areas")
            or not existing.get("templates_dir")
            or not isinstance(existing.get("next_button"), dict)
        )
        if needs_setup:
            print("No valid useless-page skip config found. Setup is required.")
            if not self.setup_useless_page_skip_config():
                print("Useless-page skip: OFF")
                return
            existing = (self.config or {}).get("useless_page_skip", {})
        elif self._ask_yes_no("Reconfigure useless-page skip now? (y/n, default n): ", default="n"):
            if not self.setup_useless_page_skip_config():
                print("Useless-page skip: OFF")
                return
            existing = (self.config or {}).get("useless_page_skip", {})

        templates_dir = str(existing.get("templates_dir", "image")).strip() or "image"
        templates = self._load_useless_skip_templates(templates_dir)
        if not templates:
            print(f"No template images found in '{templates_dir}'. Useless-page skip is disabled.")
            return

        existing["enabled"] = True
        self.config["useless_page_skip"] = existing
        Config.save(self.config)

        self.useless_skip_runtime_config = existing
        self.useless_skip_templates = templates
        self.useless_skip_enabled = True
        skip_next = existing.get("next_button", {})
        print(
            f"Useless-page skip: ON ({len(existing.get('areas', []))} areas, "
            f"{len(templates)} templates, every {self._get_useless_skip_check_interval()}s, "
            f"next=({skip_next.get('x')}, {skip_next.get('y')}))"
        )

    @staticmethod
    def _extract_white_triangle_mask(image, white_value_threshold=245, max_saturation=45):
        """Extract high-brightness, low-saturation pixels (white play icon)."""
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        lower = np.array([0, 0, int(white_value_threshold)], dtype=np.uint8)
        upper = np.array([180, int(max_saturation), 255], dtype=np.uint8)
        mask = cv2.inRange(hsv, lower, upper)

        kernel = np.ones((3, 3), np.uint8)
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=1)
        return mask

    def check_course_finished(
        self,
        iou_threshold=0.72,
        white_value_threshold=245,
        max_saturation=45,
        white_ratio_tolerance=0.15,
    ):
        """
        Return True if white triangle mask is similar to the stopped reference.

        Strategy:
        1) Binarize high-brightness white pixels (triangle) in current and reference images.
        2) Compare the masks by IoU + white-area ratio consistency.
        """
        if not self.config or "detection_area" not in self.config:
            return False

        area = self.config["detection_area"]
        ref_path = self.config.get("reference_image", "stopped_reference.png")
        if not os.path.exists(ref_path):
            return False

        screenshot = pyautogui.screenshot(
            region=(area["x1"], area["y1"], area["x2"] - area["x1"], area["y2"] - area["y1"])
        )
        current = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        reference = cv2.imread(ref_path)

        if reference is None:
            return False

        if current.shape != reference.shape:
            current = cv2.resize(current, (reference.shape[1], reference.shape[0]))

        current_mask = self._extract_white_triangle_mask(
            current,
            white_value_threshold=white_value_threshold,
            max_saturation=max_saturation,
        )
        reference_mask = self._extract_white_triangle_mask(
            reference,
            white_value_threshold=white_value_threshold,
            max_saturation=max_saturation,
        )

        total_pixels = float(current_mask.size)
        if total_pixels <= 0:
            return False

        current_white_ratio = cv2.countNonZero(current_mask) / total_pixels
        reference_white_ratio = cv2.countNonZero(reference_mask) / total_pixels
        if reference_white_ratio <= 0.01:
            return False

        intersection = cv2.countNonZero(cv2.bitwise_and(current_mask, reference_mask))
        union = cv2.countNonZero(cv2.bitwise_or(current_mask, reference_mask))
        if union == 0:
            return False

        iou = float(intersection) / float(union)
        ratio_delta = abs(current_white_ratio - reference_white_ratio)
        return iou >= float(iou_threshold) and ratio_delta <= float(white_ratio_tolerance)

    @staticmethod
    def _compute_template_similarity(current, template):
        """Return similarity score in [0, 1] using grayscale + edges."""
        if current.shape[:2] != template.shape[:2]:
            current = cv2.resize(current, (template.shape[1], template.shape[0]))

        current_gray = cv2.cvtColor(current, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        current_gray = cv2.GaussianBlur(current_gray, (3, 3), 0)
        template_gray = cv2.GaussianBlur(template_gray, (3, 3), 0)

        gray_score = float(cv2.matchTemplate(current_gray, template_gray, cv2.TM_CCOEFF_NORMED)[0][0])
        if np.isnan(gray_score):
            gray_score = 0.0

        current_edges = cv2.Canny(current_gray, 80, 160)
        template_edges = cv2.Canny(template_gray, 80, 160)
        edge_score = float(cv2.matchTemplate(current_edges, template_edges, cv2.TM_CCOEFF_NORMED)[0][0])
        if np.isnan(edge_score):
            edge_score = gray_score

        # Pixel-level fallback to stabilize low-texture scenes.
        pixel_diff = cv2.absdiff(current_gray, template_gray)
        pixel_score = 1.0 - float(np.mean(pixel_diff)) / 255.0

        combined = 0.25 * pixel_score + 0.35 * gray_score + 0.40 * edge_score
        return max(0.0, min(1.0, combined))

    def check_useless_page(self):
        """Return True if any configured area is similar to any template in templates folder."""
        if not self.useless_skip_enabled or not self.useless_skip_runtime_config:
            return False

        areas = self.useless_skip_runtime_config.get("areas", [])
        if not areas or not self.useless_skip_templates:
            return False

        threshold = float(self.useless_skip_runtime_config.get("similarity_threshold", 0.86))
        for idx, area in enumerate(areas, start=1):
            screenshot = pyautogui.screenshot(
                region=(area["x1"], area["y1"], area["x2"] - area["x1"], area["y2"] - area["y1"])
            )
            current = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            for template in self.useless_skip_templates:
                score = self._compute_template_similarity(current, template["image"])
                if score >= threshold:
                    print(
                        f"Detected useless page (area {idx}, template {template['name']}, score={score:.3f})"
                    )
                    return True
        return False

    def _start_ffmpeg_recording(self, output_file):
        """Start ffmpeg recording."""
        audio_device = self._resolve_audio_device()
        if not audio_device:
            print("No usable audio input device was found.")
            print("Run `python list_audio_devices.py` and configure `audio_device` in recorder_config.json.")
            return False

        cmd = [
            "ffmpeg",
            "-y",
            "-thread_queue_size",
            "2048",
            "-f",
            "gdigrab",
            "-framerate",
            "30",
            "-i",
            "desktop",
            "-thread_queue_size",
            "2048",
            "-f",
            "dshow",
            "-i",
            audio_device,
            "-map",
            "0:v:0",
            "-map",
            "1:a:0",
            "-c:v",
            "-preset",
            "ultrafast",
            "-crf",
            "23",
            "-pix_fmt",
            "yuv420p",
            "-r",
            "30",
            "-af",
            AUDIO_SYNC_FILTER,
            "-c:a",
            "aac",
            "-b:a",
            "192k",
            "-ar",
            "48000",
            "-ac",
            "2",
            "-max_interleave_delta",
            "0",
            output_file,
        ]

        log_file = output_file.replace(".mp4", "_ffmpeg.log")
        try:
            with open(log_file, "w", encoding="utf-8") as log:
                log.write(f"[audio_device] {audio_device}\n\n")
                self.recording_process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.DEVNULL,
                    stderr=log,
                    text=True,
                )

            time.sleep(0.8)
            if self.recording_process.poll() is None:
                self.is_recording = True
                print(f"Recording started: {output_file}")
                print(f"Audio input: {audio_device}")
                return True

            with open(log_file, "r", encoding="utf-8", errors="ignore") as log:
                print("ffmpeg failed to start.")
                print(log.read())
            return False
        except FileNotFoundError:
            print("ffmpeg not found. Please install ffmpeg and add it to PATH.")
            return False

    def _start_nvidia_recording(self, output_file):
        """Toggle NVIDIA recording on using the configured hotkey."""
        hotkey = self._get_nvidia_hotkey()
        try:
            pyautogui.hotkey(*hotkey)
            time.sleep(1.0)
        except Exception as exc:
            print(f"Failed to trigger NVIDIA start hotkey ({'+'.join(hotkey)}): {exc}")
            return False

        self.is_recording = True
        self.current_output_file = output_file
        print(f"NVIDIA recording toggled on via {'+'.join(hotkey)}")
        print("Output file will be created by NVIDIA app in its configured save folder.")
        return True

    def start_recording(self, output_file):
        """Start recording based on configured recording mode."""
        if self._get_recording_mode() == "nvidia":
            return self._start_nvidia_recording(output_file)
        return self._start_ffmpeg_recording(output_file)

    def _stop_ffmpeg_recording(self, current_file=None):
        """Stop ffmpeg recording."""
        if not self.recording_process:
            return

        print("Stopping recording...")
        try:
            self.recording_process.communicate(input="q\n", timeout=10)
        except subprocess.TimeoutExpired:
            self.recording_process.terminate()
            try:
                self.recording_process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.recording_process.kill()
                self.recording_process.wait()
        except Exception as exc:
            print(f"Stop error: {exc}")

        self.recording_process = None
        self.is_recording = False

        if current_file and os.path.exists(current_file):
            size_mb = os.path.getsize(current_file) / (1024 * 1024)
            print(f"Saved: {current_file} ({size_mb:.2f} MB)")

    def _stop_nvidia_recording(self):
        """Toggle NVIDIA recording off using the configured hotkey."""
        hotkey = self._get_nvidia_hotkey()
        print("Stopping recording...")
        try:
            pyautogui.hotkey(*hotkey)
            time.sleep(1.0)
        except Exception as exc:
            print(f"Failed to trigger NVIDIA stop hotkey ({'+'.join(hotkey)}): {exc}")
            return

        self.recording_process = None
        self.is_recording = False
        print(f"NVIDIA recording toggled off via {'+'.join(hotkey)}")

    def stop_recording(self, current_file=None):
        """Stop recording based on configured recording mode."""
        if self._get_recording_mode() == "nvidia":
            self._stop_nvidia_recording()
            return
        self._stop_ffmpeg_recording(current_file)

    def _get_safe_mouse_position(self):
        """Pick a corner position that is outside the detection area."""
        width, height = pyautogui.size()
        margin = 40
        right = max(margin, width - margin)
        bottom = max(margin, height - margin)
        candidates = [
            (right, margin),
            (right, bottom),
            (margin, bottom),
            (margin, margin),
        ]

        all_areas = []
        main_area = (self.config or {}).get("detection_area")
        if isinstance(main_area, dict):
            all_areas.append(main_area)
        skip_cfg = (self.config or {}).get("useless_page_skip", {})
        for area in skip_cfg.get("areas", []):
            if isinstance(area, dict):
                all_areas.append(area)
        if not all_areas:
            return candidates[0]

        padding = 20

        def in_detection_area(x, y):
            for area in all_areas:
                if (
                    area["x1"] - padding <= x <= area["x2"] + padding
                    and area["y1"] - padding <= y <= area["y2"] + padding
                ):
                    return True
            return False

        for candidate in candidates:
            if not in_detection_area(*candidate):
                return candidate
        return candidates[0]

    def _move_mouse_to_safe_area(self):
        """Move mouse away from player controls so it won't affect visual detection."""
        try:
            x, y = self._get_safe_mouse_position()
            pyautogui.moveTo(x, y, duration=0.15)
            print(f"Moved mouse to safe area: ({x}, {y})")
        except Exception as exc:
            print(f"Failed to move mouse to safe area: {exc}")

    def click_next_button(self):
        if not self.config or "next_button" not in self.config:
            return False
        button = self.config["next_button"]
        pyautogui.click(button["x"], button["y"])
        print(f"Clicked Next at ({button['x']}, {button['y']})")
        self._move_mouse_to_safe_area()
        time.sleep(2)
        return True

    def click_useless_skip_next_button(self):
        cfg = self.useless_skip_runtime_config or {}
        button = cfg.get("next_button")
        if not isinstance(button, dict) or "x" not in button or "y" not in button:
            return False
        pyautogui.click(button["x"], button["y"])
        print(f"Clicked Useless-Skip Next at ({button['x']}, {button['y']})")
        self._move_mouse_to_safe_area()
        time.sleep(2)
        return True

    def click_play_button(self):
        if not self.config or "play_button" not in self.config:
            return False
        button = self.config["play_button"]
        pyautogui.click(button["x"], button["y"])
        print(f"Clicked Play at ({button['x']}, {button['y']})")
        self._move_mouse_to_safe_area()
        time.sleep(1)
        return True

    @staticmethod
    def _build_output_file(index, output_dir):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(output_dir, f"course_{index}_{timestamp}.mp4")

    @staticmethod
    def _rename_as_timeout_video(path):
        if not os.path.exists(path):
            return
        base = os.path.basename(path).replace(".mp4", "")
        new_name = f"{base}_timeout.mp4"
        new_path = os.path.join(os.path.dirname(path), new_name)
        shutil.move(path, new_path)
        print(f"Renamed timeout file: {new_name}")

    def monitor_and_record(self):
        """Main loop."""
        if not self.config:
            print("No config loaded.")
            return

        os.makedirs("recordings", exist_ok=True)
        output_dir = "recordings"
        max_recording_duration = 60 * 60
        check_interval = 5
        useless_check_interval = self._get_useless_skip_check_interval()

        print("\n" + "=" * 60)
        print("Auto recording started. Press Ctrl+C to stop.")
        print("=" * 60)

        try:
            while True:
                if self.playing_state == 0:
                    print("State: waiting for play. Trying to click Play...")
                    if not self.click_play_button():
                        print("Cannot click Play. Stop.")
                        return
                    self.playing_state = 1
                    time.sleep(2)

                output_file = self._build_output_file(self.current_file_number, output_dir)
                self.current_output_file = output_file
                print(f"\n[Lesson {self.current_file_number}] {output_file}")

                if not self.start_recording(output_file):
                    return

                start_time = time.time()
                checks = 0
                last_useless_check_at = start_time

                while True:
                    time.sleep(check_interval)
                    checks += 1
                    elapsed = time.time() - start_time

                    if checks % 6 == 0:
                        print(f"Recording... {elapsed / 60:.1f} min")

                    if elapsed >= max_recording_duration:
                        print("Timeout reached (60 min).")
                        self.stop_recording(output_file)
                        if self._get_recording_mode() == "ffmpeg":
                            self._rename_as_timeout_video(output_file)
                        else:
                            print("NVIDIA mode: timeout clip remains in NVIDIA output folder.")
                        return

                    now = time.time()
                    if self.useless_skip_enabled and now - last_useless_check_at >= useless_check_interval:
                        last_useless_check_at = now
                        if self.check_useless_page():
                            print("Detected useless page. Skip to next lesson.")
                            if not self.click_useless_skip_next_button():
                                print("Cannot click useless-skip Next. Stop.")
                                return
                            self.playing_state = 0
                            print("State machine reset: playing_state=0")
                            if not self.click_play_button():
                                print("Cannot click Play after useless-page skip. Stop.")
                                return
                            self.playing_state = 1
                            print("State machine update: playing_state=1")
                            continue

                    if self.check_course_finished():
                        time.sleep(2)
                        if self.check_course_finished():
                            print("Detected lesson finished.")
                            self.stop_recording(output_file)

                            if not self.click_next_button():
                                print("Cannot click Next. Stop.")
                                return

                            self.playing_state = 0
                            self.current_file_number += 1
                            time.sleep(3)
                            break

        except KeyboardInterrupt:
            print("\nInterrupted by user.")
            if self.is_recording:
                self.stop_recording(self.current_output_file)

    def run(self):
        print("Course Auto Recorder")
        print("=" * 60)

        self.config = Config.load()
        if not self.config:
            print("No config found. Starting setup...")
            if not self.setup_config():
                return
        else:
            print("Loaded existing config.")
            print(f"Next button: {self.config.get('next_button')}")
            print(f"Play button: {self.config.get('play_button')}")
            if not self.config.get("recording_mode"):
                self.config["recording_mode"] = DEFAULT_RECORDING_MODE
                Config.save(self.config)
            if not self.config.get("audio_device"):
                resolved_audio_device = self._resolve_audio_device()
                if resolved_audio_device:
                    self.config["audio_device"] = resolved_audio_device
                    Config.save(self.config)
            print(f"Recording mode: {self._get_recording_mode()}")
            print(f"Audio device: {self.config.get('audio_device', '(not set)')}")
            print(f"NVIDIA toggle hotkey: {'+'.join(self._get_nvidia_hotkey())}")
            print("Reconfigure? (y/n): ", end="")
            if input().strip().lower() == "y":
                if not self.setup_config():
                    return

        print(f"Start file number (current: {self.current_file_number}): ", end="")
        answer = input().strip()
        if answer.isdigit():
            self.current_file_number = int(answer)

        self._prepare_useless_page_skip()
        self.monitor_and_record()


def main():
    recorder = CourseRecorder()
    recorder.run()


if __name__ == "__main__":
    main()
