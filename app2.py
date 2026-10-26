# =====================================================================
# 1. 기본 설정 및 라이브러리 임포트
# =====================================================================
from flask import Flask, render_template, request, jsonify
import pyautogui
import time
from datetime import datetime, timedelta
import logging
import pandas as pd
import os
import ctypes
from werkzeug.utils import secure_filename
import traceback
from PIL import Image
import re
import pytesseract
import pyperclip
import threading
import uiautomation as auto
import comtypes
from logging.handlers import TimedRotatingFileHandler
import sys
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # 관리자 권한이 없으면 UAC 창을 띄우며 스스로를 관리자 권한으로 다시 실행
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(f'"{arg}"' for arg in sys.argv), None, 1
    )
    sys.exit()

# uiautomation 전역 검색 타임아웃을 2초로 제한
auto.SetGlobalSearchTimeout(2)

# Tesseract 설치 경로 지정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB 최대 파일 크기

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# PyAutoGUI 설정
pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = True

# 로깅 설정 (터미널 + 날짜별 파일 저장)
_log_formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')

_stream_handler = logging.StreamHandler()
_stream_handler.setFormatter(_log_formatter)

_file_handler = TimedRotatingFileHandler(
    'ocr_log.txt',
    when='midnight',
    interval=30,
    backupCount=12,
    encoding='utf-8'
)
_file_handler.setFormatter(_log_formatter)
_file_handler.suffix = '%Y%m'

_root_logger = logging.getLogger()
_root_logger.setLevel(logging.INFO)
_root_logger.addHandler(_stream_handler)
_root_logger.addHandler(_file_handler)

logger = logging.getLogger(__name__)


# =====================================================================
# 2. 전역 변수 및 예외 정의
# =====================================================================

# 희망일자 체크 시 적용할 '분류' 값
ALLOWED_CLASSIFICATION = frozenset({'D52', 'D53', 'D54'})

# 긴급 정지 플래그 및 예외
_stop_flag = False

class _StopAutomation(BaseException):
    """ESC 중단 신호 — BaseException 상속으로 어떤 except Exception에도 걸리지 않음"""
    pass

# 진행 상황 저장 전역 딕셔너리
progress_state = {
    'is_running': False,
    'total': 0,
    'current': 0,
    'results': []
}

def _trigger_stop():
    global _stop_flag
    _stop_flag = True
    logger.warning("ESC 감지! 자동화를 중단합니다...")

def reset_stop_flag():
    global _stop_flag
    _stop_flag = False

def check_stop():
    if _stop_flag:
        raise _StopAutomation("사용자가 ESC로 자동화를 중단했습니다.")

def _start_stop_listener():
    try:
        import keyboard
        keyboard.add_hotkey("esc", _trigger_stop)
        logger.info("긴급정지 리스너 시작 (ESC키로 중단 가능)")
        keyboard.wait()
    except ImportError:
        logger.warning("keyboard 모듈 없음. pip install keyboard 로 설치하세요.")
    except Exception as e:
        logger.warning(f"긴급정지 리스너 오류: {e}")

_listener_thread = threading.Thread(target=_start_stop_listener, daemon=True)
_listener_thread.start()


# =====================================================================
# 3. UI Automation 및 창 제어 헬퍼
# =====================================================================

class RECT(ctypes.Structure):
    _fields_ = [
        ("left",   ctypes.c_long),
        ("top",    ctypes.c_long),
        ("right",  ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]

_vsp_ctrl_cache = None
_amis_hwnd_cache = None
_vsp_resized = False

def _get_vsp_window_cached():
    """채혈접수 창을 캐시로 반환 (첫 번째만 탐색, 이후 재사용)"""
    global _vsp_ctrl_cache
    comtypes.CoInitialize()
    if _vsp_ctrl_cache is not None:
        try:
            _ = _vsp_ctrl_cache.Name
            return _vsp_ctrl_cache
        except:
            _vsp_ctrl_cache = None
    try:
        amis = auto.WindowControl(Name='AMIS3', timeout=2)
        vsp  = auto.WindowControl(
            searchFromControl=amis,
            Name='채혈접수[VSPSSSP010S]_서관2층검사실',
            timeout=2
        )
        _vsp_ctrl_cache = vsp
        logger.info("채혈접수 창 캐시 저장")
        return vsp
    except Exception as e:
        logger.warning(f"채혈접수 창 찾기 실패: {e}")
        return None


def ui_input_hospital_id(hospital_id):
    """uiautomation으로 병원등록번호 입력"""
    try:
        comtypes.CoInitialize()
        vsp = _get_vsp_window_cached()
        if not vsp:
            return False
        panel = auto.PaneControl(searchFromControl=vsp, Name='aPanelControl1', timeout=1)
        for ctrl in panel.GetChildren():
            if ctrl.ControlTypeName == 'PaneControl':
                for child in ctrl.GetChildren():
                    if child.ControlTypeName == 'EditControl':
                        try:
                            val = child.GetValuePattern().Value or ''
                            if val.isdigit() or val == '':
                                child.SetFocus()
                                child.GetValuePattern().SetValue(hospital_id)
                                time.sleep(0.05)
                                logger.info(f"uiautomation 등록번호 입력 성공: {hospital_id}")
                                return True
                        except:
                            continue
        return False
    except Exception as e:
        logger.warning(f"uiautomation 등록번호 입력 실패 (좌표 방식으로 fallback): {e}")
        return False


def ui_get_patient_name():
    """uiautomation으로 환자명 읽기"""
    try:
        comtypes.CoInitialize()
        vsp = _get_vsp_window_cached()
        if not vsp:
            return ''

        korean_pattern = re.compile(r'^[가-힣]{2,5}$')
        panel = auto.PaneControl(searchFromControl=vsp, Name='aPanelControl1', timeout=1)
        
        for ctrl, depth in auto.WalkControl(panel, maxDepth=3):
            if ctrl.ControlTypeName in ['EditControl', 'TextControl', 'CustomControl']:
                try:
                    val = ctrl.GetValuePattern().Value or ''
                    val = val.strip()
                    if korean_pattern.match(val):
                        logger.info(f"uiautomation 환자명 발견(Value): {val}")
                        return val
                except:
                    pass

                name = ctrl.Name or ''
                name = name.strip()
                if korean_pattern.match(name):
                    logger.info(f"uiautomation 환자명 발견(Name): {name}")
                    return name

        return ''
    except Exception as e:
        logger.debug(f"uiautomation 환자명 읽기 실패 (클립보드로 넘어갑니다): {e}")
        return ''


def dismiss_amis_popup(appear_timeout=0.35, settle_timeout=0.4,
                       poll_interval=0.05, max_dismiss=6, hard_cap=3.0):
    """'AMIS 메시지' 팝업이 나타나면 '확인' 버튼을 찾아 닫는다."""
    import win32gui

    _found = []

    def _enum_cb(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and 'AMIS 메시지' in title:
                    _found.append(hwnd)
        except Exception:
            pass

    dismissed_any = False
    dismiss_count = 0
    hard_deadline = time.perf_counter() + hard_cap
    deadline = time.perf_counter() + appear_timeout

    while time.perf_counter() < deadline and time.perf_counter() < hard_deadline:
        _found.clear()
        try:
            win32gui.EnumWindows(_enum_cb, None)
        except Exception as e:
            logger.warning(f"팝업 창 열거 실패: {e}")
            time.sleep(poll_interval)
            continue

        if not _found:
            time.sleep(poll_interval)
            continue

        hwnd = _found[0]
        logger.info(f"팝업 감지(win32): hwnd={hwnd} '{win32gui.GetWindowText(hwnd)}'")
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        time.sleep(0.05)
        pyautogui.press('enter')
        logger.info("팝업 Enter 전송 완료")

        dismissed_any = True
        dismiss_count += 1
        if dismiss_count >= max_dismiss:
            logger.warning(f"팝업을 {max_dismiss}회 닫았는데도 계속 감지됨 → 중단")
            break

        time.sleep(0.2)
        deadline = time.perf_counter() + settle_timeout

    return dismissed_any


def ui_click_action_and_register():
    """uiautomation으로 Action 버튼 클릭 후 채혈접수(F4) 클릭"""
    try:
        comtypes.CoInitialize()
        vsp = _get_vsp_window_cached()
        if not vsp:
            return False

        action_btn = None
        try:
            scroll_panel = auto.PaneControl(searchFromControl=vsp, Name='scrollPanel', timeout=1)
            panel2 = auto.PaneControl(searchFromControl=scroll_panel, Name='aPanelControl2', timeout=1)
            action_btn = auto.ButtonControl(searchFromControl=panel2, Name='Action', timeout=1)
            logger.info("aPanelControl2에서 Action 버튼 발견")
        except Exception as e:
            logger.warning(f"aPanelControl2 Action 탐색 실패: {e}, 전체 탐색 시도")
            def find_action_in_vsp(ctrl, depth=0):
                if depth > 10:
                    return None
                if ctrl.ControlTypeName == 'ButtonControl' and ctrl.Name == 'Action':
                    return ctrl
                for child in ctrl.GetChildren():
                    result = find_action_in_vsp(child, depth + 1)
                    if result:
                        return result
                return None
            action_btn = find_action_in_vsp(vsp)

        if not action_btn:
            logger.warning("Action 버튼을 찾지 못했습니다.")
            return False

        logger.info("Action 버튼 클릭")
        try:
            action_btn.GetInvokePattern().Invoke()
        except Exception as e:
            logger.warning(f"Invoke 실패, 일반 클릭 시도: {e}")
            action_btn.Click()
            
        time.sleep(0.3)
        import pyautogui as _pag
        _pag.press('f4')
        logger.info("채혈접수 F4 키 입력 완료")
        return True

    except Exception as e:
        logger.warning(f"uiautomation Action 클릭 실패: {e}")
        return False


def find_window_by_title(title):
    """윈도우 제목으로 창을 찾아 활성화하고 작업표시줄 제외 최대화"""
    try:
        import win32gui
        import win32con
        import win32process
        import win32api

        found_hwnd = []

        def enum_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if window_text and title in window_text:
                    found_hwnd.append(hwnd)

        win32gui.EnumWindows(enum_callback, None)

        if not found_hwnd:
            logger.warning(f"'{title}' 창을 찾지 못했습니다.")
            return False

        hwnd = found_hwnd[0]
        logger.info(f"창 발견: {win32gui.GetWindowText(hwnd)}")

        if win32gui.IsIconic(hwnd):
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                print(f"오류: {e}")
            time.sleep(0.2)
        check_stop()

        current_hwnd = win32gui.GetForegroundWindow()
        current_thread_id = win32api.GetCurrentThreadId()
        target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
        foreground_thread_id, _ = win32process.GetWindowThreadProcessId(current_hwnd)

        ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, True)
        ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, True)

        try: 
            import win32con 
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                logger.warning(f"SetForegroundWindow 실패: {e}")
                
            ctypes.windll.user32.SetFocus(hwnd)
            ctypes.windll.user32.SetActiveWindow(hwnd)

            work_area = RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(work_area), 0)
            
            total_width = work_area.right - work_area.left
            target_width = (total_width * 2) // 3
            height = work_area.bottom - work_area.top
            
            try:
                win32gui.MoveWindow(hwnd, work_area.left, work_area.top, target_width, height, True)
                logger.info(f"창 크기 조정 완료: 좌측 2/3 배치 ({target_width}px)")
            except Exception as e:
                logger.warning(f"창 이동 실패: {e}")

            time.sleep(0.05)

            ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, False)
            ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, False)

            time.sleep(0.1)
            active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            logger.info(f"활성화된 창 제목: {active_title}")

            return True

        except Exception as e:
            logger.error(f"창 활성화 프로세스 중 오류(무시): {e}")
            try:
                ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, False)
                ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, False)
            except:
                pass
            
            try:
                import pygetwindow as gw
                all_windows = []
                try:
                    all_windows = gw.getAllWindows()
                except AttributeError:
                    all_windows = gw.getWindows()

                for w in all_windows:
                    if w.title and title in w.title:
                        if not w.isActive:
                            w.activate()
                        logger.info(f"창 찾음 (pygetwindow): {w.title}")
                        return True
            except Exception as ex:
                logger.warning(f"pygetwindow 최종 시도 실패: {ex}")
            
            return True

    except Exception as e:
        logger.error(f"find_window_by_title 전체 오류: {e}")
        return False


def find_amis_window():
    """AMIS 창을 찾아 활성화 (hwnd 캐싱)"""
    global _amis_hwnd_cache
    import win32gui
    import win32con

    if _amis_hwnd_cache is not None:
        try:
            if win32gui.IsWindow(_amis_hwnd_cache) and win32gui.IsWindowVisible(_amis_hwnd_cache):
                active_hwnd = win32gui.GetForegroundWindow()
                if active_hwnd != _amis_hwnd_cache:
                    if win32gui.IsIconic(_amis_hwnd_cache):
                        win32gui.ShowWindow(_amis_hwnd_cache, win32con.SW_RESTORE)
                        time.sleep(0.1)
                    try:
                        win32gui.SetForegroundWindow(_amis_hwnd_cache)
                    except:
                        ctypes.windll.user32.SetFocus(_amis_hwnd_cache)
                        ctypes.windll.user32.SetActiveWindow(_amis_hwnd_cache)
                logger.info(f"[캐시] AMIS 창 재사용 (hwnd={_amis_hwnd_cache})")
                return True
        except Exception as e:
            logger.warning(f"캐시된 hwnd 사용 실패, 재탐색: {e}")
            _amis_hwnd_cache = None

    search_terms = ["AMIS", "VSPSSSP010S", "채혈접수", "AMIS 3.0"]
    for term in search_terms:
        try:
            if find_window_by_title(term):
                import win32gui as _wg
                found = []
                def _cb(hwnd, _):
                    if _wg.IsWindowVisible(hwnd) and term in (_wg.GetWindowText(hwnd) or ''):
                        found.append(hwnd)
                _wg.EnumWindows(_cb, None)
                if found:
                    _amis_hwnd_cache = found[0]
                    logger.info(f"AMIS 창 hwnd 캐시 저장: {_amis_hwnd_cache}")
                return True
        except Exception as e:
            logger.warning(f"'{term}' 검색 중 오류: {e}")
            continue
    return False


def reset_vsp_resize():
    global _vsp_resized
    _vsp_resized = False


def resize_vsp_window():
    """채혈접수[VSPSSSP010S] MDI 내부 창 크기 조정"""
    global _vsp_resized
    if _vsp_resized:
        return

    try:
        comtypes.CoInitialize()
        amis = auto.WindowControl(Name='AMIS3', timeout=2)
        vsp  = auto.WindowControl(
            searchFromControl=amis,
            Name='채혈접수[VSPSSSP010S]_서관2층검사실',
            timeout=2
        )

        rect = vsp.BoundingRectangle
        cur_left   = rect.left
        cur_top    = rect.top
        cur_width  = rect.width()
        cur_height = rect.height()

        TARGET_HEIGHT = max(cur_height, 840)
        TARGET_WIDTH  = cur_width + 20

        try:
            tp = vsp.GetTransformPattern()
            if tp.CanResize:
                tp.Resize(TARGET_WIDTH, TARGET_HEIGHT)
                logger.info(f"채혈접수 창 크기 조정: {cur_width}x{cur_height} → {TARGET_WIDTH}x{TARGET_HEIGHT}")
            if tp.CanMove:
                tp.Move(cur_left, cur_top)
        except Exception as e:
            logger.warning(f"TransformPattern 크기 조정 실패: {e}")

        _vsp_resized = True
        time.sleep(0.2)
    except Exception as e:
        logger.warning(f"채혈접수 창 크기 조정 실패 (무시하고 계속): {e}")
        _vsp_resized = True


# =====================================================================
# 4. OCR 및 유틸리티 함수
# =====================================================================

def allowed_file(filename):
    """허용된 파일 확장자인지 확인"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def preprocess_for_ocr(img, binarize=True):
    """강화된 OCR 전처리: 4배 확대 + 자동대비 + 이진화"""
    from PIL import ImageOps
    img = img.resize((img.width * 4, img.height * 4), Image.Resampling.LANCZOS)
    img = img.convert('L')
    img = ImageOps.autocontrast(img)
    if binarize:
        img = img.point(lambda x: 0 if x < 140 else 255, '1')
        img = img.convert('L')
    return img


def extract_hospital_id(text):
    """텍스트 내에서 연속된 8자리 숫자를 모두 찾아 리스트로 반환"""
    pattern = r'\b\d{8}\b'
    matches = re.findall(pattern, text)
    return matches


# =====================================================================
# 5. 자동화 핵심 로직 (날짜 체크 및 등록)
# =====================================================================

def click_date_checkboxes_in_range():
    """희망일자(오늘±3일) + 분류(D52/D53/D54) 조건 충족 행을 uiautomation으로 클릭"""
    today = datetime.now().date()

    valid_dates = set()
    for delta in range(-3, 4):
        d = today + timedelta(days=delta)
        valid_dates.add(d.strftime('%Y%m%d'))
        valid_dates.add(d.strftime('%Y-%m-%d'))
        valid_dates.add(d.strftime('%Y/%m/%d'))
        valid_dates.add(d.strftime('%y-%m-%d'))
        valid_dates.add(d.strftime('%y/%m/%d'))
        valid_dates.add(d.strftime('%m/%d/%Y'))

    logger.info(f"유효 날짜 (오늘 기준 샘플): {today.strftime('%Y%m%d')} / {today.strftime('%Y-%m-%d')}")

    clicked_count = 0
    clicked_specimens = set()

    auto.SetGlobalSearchTimeout(0.3)

    try:
        t0 = time.perf_counter()
        comtypes.CoInitialize()

        amis = auto.WindowControl(Name='AMIS3', timeout=0.5)
        vsp = auto.WindowControl(searchFromControl=amis,
                                 Name='채혈접수[VSPSSSP010S]_서관2층검사실', timeout=0.5)
        scroll_panel = auto.PaneControl(searchFromControl=vsp, Name='scrollPanel', timeout=1)

        target_table = None
        try:
            children = scroll_panel.GetChildren()
            for child in children:
                if child.ControlTypeName == 'TableControl':
                    target_table = child
                    break
                if child.ControlTypeName == 'PaneControl':
                    try:
                        for grandchild in child.GetChildren():
                            if grandchild.ControlTypeName == 'TableControl':
                                target_table = grandchild
                                break
                    except:
                        pass
                if target_table:
                    break
        except Exception as e:
            logger.warning(f"테이블 직접 탐색 실패, 재귀로 fallback: {e}")

        if not target_table:
            def find_target_table(ctrl, depth=0):
                if depth > 5:
                    return None
                if ctrl.ControlTypeName == 'TableControl':
                    try:
                        for child in ctrl.GetChildren():
                            if child.ControlTypeName == 'CustomControl' and child.Name == '헤더 패널':
                                headers = [h.Name for h in child.GetChildren()]
                                if '분류' in headers and '희망일자' in headers:
                                    return ctrl
                    except:
                        pass
                try:
                    for child in ctrl.GetChildren():
                        result = find_target_table(child, depth + 1)
                        if result:
                            return result
                except:
                    pass
                return None
            target_table = find_target_table(scroll_panel)

        if not target_table:
            logger.error("처방목록 테이블을 찾을 수 없습니다.")
            return 0

        try:
            data_panel = auto.CustomControl(searchFromControl=target_table, Name='데이터 패널', timeout=1)
        except Exception:
            logger.info("⚡ 데이터 패널 없음 → 처방목록 없음, 즉시 반환")
            return 0

        try:
            rect = data_panel.BoundingRectangle
            center_x = rect.left + (rect.width() // 2)
            center_y = rect.top + (rect.height() // 2)
            pyautogui.moveTo(center_x, center_y, duration=0)
            pyautogui.scroll(30)
            time.sleep(0.3)
            logger.info("⚡ 처방목록 스크롤 초기화 (맨 위로 이동)")
        except Exception as e:
            logger.warning(f"스크롤 초기화 실패: {e}")

        seen_row_names = set()
        empty_streak   = 0
        scroll_attempts = 0
        MAX_SCROLL = 10
        prev_visible_names = None
        _loop_deadline = time.perf_counter() + 15

        while scroll_attempts < MAX_SCROLL:
            check_stop()
            if time.perf_counter() > _loop_deadline:
                logger.warning("⚠️ 처방목록 탐색 15초 초과 → 다음 환자로 진행")
                break
            rows = data_panel.GetChildren()

            if not rows:
                logger.info("⚡ 처방목록 없음 → 즉시 다음 환자로")
                break

            current_visible_names = set(r.Name for r in rows)
            if prev_visible_names is not None and current_visible_names == prev_visible_names:
                logger.info("⚡ 스크롤 후 동일한 행 목록 → 리스트 끝 도달, 탐색 종료")
                break
            prev_visible_names = current_visible_names

            new_rows_found = False

            for row in rows:
                check_stop()
                try:
                    row_name = row.Name
                    if row_name in seen_row_names:
                        continue
                    seen_row_names.add(row_name)
                    new_rows_found = True

                    items = row.GetChildren()
                    date_val      = None
                    class_val     = None
                    has_any_value = False
                    specimen_item = None
                    specimen_no   = None
                    all_col_log   = []

                    for item in items:
                        check_stop()
                        item_name  = item.Name or ''
                        item_value = ''
                        try:
                            item_value = item.GetValuePattern().Value or ''
                        except:
                            pass
                        v = item_value.strip()

                        if v:
                            all_col_log.append(f"{item_name}={v}")
                            has_any_value = True

                        if item_name.startswith('희망일자 row'):
                            date_val = v
                        elif item_name.startswith('분류 row'):
                            class_val = v
                        elif item_name.startswith('검체번호 row'):
                            if v and len(v) >= 10 and v.isdigit():
                                specimen_no   = v
                                specimen_item = item

                        if date_val is not None and class_val is not None and specimen_no is not None:
                            break

                    if not has_any_value:
                        empty_streak += 1
                        if empty_streak >= 1:
                            new_rows_found = False
                            break
                        continue
                    else:
                        empty_streak = 0

                    is_date_ok  = date_val  in valid_dates
                    is_class_ok = class_val in ALLOWED_CLASSIFICATION

                    if is_date_ok and is_class_ok:
                        if specimen_item is None:
                            continue
                        if specimen_no in clicked_specimens:
                            continue

                        logger.info(f"🎯 조건 충족! {row_name} 클릭 (날짜={date_val}, 분류={class_val}, 검체번호={specimen_no})")
                        specimen_item.Click()
                        clicked_specimens.add(specimen_no)
                        clicked_count += 1
                        time.sleep(0.05)

                except Exception as e:
                    logger.debug(f"행 처리 중 오류: {e}")
                    continue

            if not new_rows_found:
                break

            scroll_attempts += 1
            try:
                rect = data_panel.BoundingRectangle
                center_x = rect.left + (rect.right - rect.left) // 2
                center_y = rect.top  + (rect.bottom - rect.top) // 2
                pyautogui.moveTo(center_x, center_y, duration=0)
                pyautogui.scroll(-5)
                time.sleep(0.1)
            except Exception as e:
                logger.warning(f"스크롤 실패: {e}")
                break

    except _StopAutomation:
        raise
    except Exception as e:
        logger.error(f"uiautomation 테이블 읽기 실패: {e}")
        logger.error(traceback.format_exc())
        return 0
    finally:
        auto.SetGlobalSearchTimeout(1)

    logger.info(f"✅ 클릭 완료: 총 {clicked_count}개 행")
    return clicked_count


def diagnose_coordinates():
    """AMIS 화면에서 좌표 확인용 진단 함수"""
    print("5초 뒤 캡처 시작... AMIS 채혈접수 화면으로 전환하세요.")
    time.sleep(5)

    test_y = 200
    test_h = 20

    print("\n[날짜 컬럼 X 탐색 (50~520, 15px 간격)]")
    for test_x in range(50, 520, 15):
        ss = pyautogui.screenshot(region=(test_x, test_y, 140, test_h))
        raw = pytesseract.image_to_string(
            preprocess_for_ocr(ss, binarize=False),
            config='--psm 7 -c tessedit_char_whitelist=0123456789-/ '
        ).strip()
        m = re.search(r'(\d{4})[-/]?(\d{2})[-/]?(\d{2})', raw)
        marker = f"✅ → {''.join(m.groups())}" if m else ''
        if raw:
            print(f"  x={test_x:3d}: '{raw}' {marker}")

    print("\n[분류 컬럼 X 탐색 (50~620, 15px 간격)]")
    for test_x in range(50, 620, 15):
        ss = pyautogui.screenshot(region=(test_x, test_y, 90, test_h))
        raw = pytesseract.image_to_string(
            preprocess_for_ocr(ss, binarize=False),
            config='--psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 '
        ).strip().upper()
        clean = re.sub(r'[^A-Z0-9]', '', raw)
        found = next((c for c in ALLOWED_CLASSIFICATION if c in clean), None)
        marker = f"✅ → {found}" if found else ''
        if raw:
            print(f"  x={test_x:3d}: '{raw}' {marker}")

    print("\n=== 진단 완료 ===")


def automate_registration(hospital_id, patient_name=''):
    """등록번호를 받아서 자동화 작업을 수행하는 메인 함수"""
    try:
        logger.info(f"자동화 시작: 등록번호 {hospital_id}")

        logger.info("1단계: AMIS 3.0 창으로 이동")
        try:
            window_found = find_amis_window()
            if not window_found:
                error_msg = "AMIS 3.0 창을 찾을 수 없습니다."
                logger.warning(error_msg)
                return (False, error_msg, patient_name)
        except Exception as e:
            error_msg = f"창 찾기 중 오류: {str(e)}"
            logger.error(error_msg)
            return (False, error_msg, patient_name)

        check_stop()
        resize_vsp_window()

        registration_input_x = 144
        registration_input_y = 89
        logger.info(f"2단계: 등록번호 입력: {hospital_id}")

        if not ui_input_hospital_id(hospital_id):
            try:
                pyautogui.click(registration_input_x, registration_input_y)
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.write(hospital_id, interval=0.03)
                time.sleep(0.1)
            except Exception as e:
                error_msg = f"등록번호 입력 실패: {str(e)}"
                logger.error(error_msg)
                return (False, error_msg, patient_name)

        pyautogui.press('enter')
        check_stop()

        time.sleep(0.1)
        _t_pop = time.perf_counter()
        if dismiss_amis_popup():
            logger.info("팝업창 닫기 완료")
            time.sleep(0.1)
        check_stop()

        _t_vsp = time.perf_counter()
        _vsp_for_wait = _get_vsp_window_cached()
        check_stop()

        _wait_start = time.perf_counter()
        if _vsp_for_wait:
            for _ in range(20):
                check_stop()
                time.sleep(0.1)
                try:
                    _sp_chk = auto.PaneControl(
                        searchFromControl=_vsp_for_wait, Name='scrollPanel')
                    if _sp_chk.Exists(0, 0) and _sp_chk.GetChildren():
                        break
                except:
                    pass
        else:
            time.sleep(0.5)

        if not patient_name or patient_name == "이름 없음" or patient_name == "조회 중":
            logger.info("4.5단계: 환자명 읽기 시도")
            patient_name = ui_get_patient_name()
            
            if not patient_name:
                try:
                    NAME_X, NAME_Y = 275, 90
                    pyautogui.doubleClick(NAME_X, NAME_Y)
                    time.sleep(0.05)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(0.05)
                    patient_name = pyperclip.paste().strip()
                    
                    if not re.match(r'^[가-힣]{2,5}$', patient_name):
                        patient_name = "확인불가"
                except Exception as e:
                    logger.warning(f"환자명 추출 최종 실패: {e}")
                    patient_name = "확인불가"

        check_stop()

        logger.info("5단계: 희망일자 확인 및 체크박스 클릭")
        try:
            clicked_count = click_date_checkboxes_in_range()

            if clicked_count > 0:
                logger.info(f"{clicked_count}개의 조건을 충족하는 행을 클릭했습니다.")

                if not ui_click_action_and_register():
                    pyautogui.click(2510, 144)
                    time.sleep(0.3)
                    pyautogui.click(2510, 167)

                if dismiss_amis_popup():
                    logger.info("채혈접수 등록 후 팝업창 닫기 완료")

                return (True, "바코드를 출력하였습니다.", patient_name)

            else:
                logger.warning("충족되는 행이 없습니다.")
                try:
                    pyautogui.click(registration_input_x, registration_input_y)
                    time.sleep(0.05)
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('delete')
                    logger.info("입력란 클리어 완료")
                except Exception as e:
                    logger.warning(f"입력란 클리어 실패(무시): {e}")
                return (True, "충족되는 행이 없습니다.", patient_name)

        except _StopAutomation:
            raise
        except Exception as e:
            error_msg = f"날짜 체크박스 클릭 중 오류: {str(e)}"
            logger.error(error_msg)
            return (False, error_msg, patient_name)

    except _StopAutomation as e:
        logger.warning(str(e))
        return (False, str(e), patient_name)
    except Exception as e:
        error_msg = f"자동화 프로세스 중 오류: {str(e)}"
        logger.error(error_msg)
        return (False, error_msg, patient_name)


# =====================================================================
# 6. Flask 라우트 및 비동기 백그라운드 처리
# =====================================================================

@app.route('/')
def index():
    return render_template('index.html')


def process_excel_background(df, filepath):
    global progress_state
    try:
        hospital_col = df.columns[0]
        patient_col  = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        logger.info(f"엑셀 컬럼 매핑 — 등록번호: [{hospital_col}], 환자명: [{patient_col}]")

        for index, row in df.iterrows():
            if _stop_flag:
                progress_state['results'].insert(0, {'id': '-', 'name': '-', 'status': '🚨 ESC 키에 의해 작업이 중단되었습니다.'})
                break

            patient_name = str(row[patient_col]).strip() if pd.notna(row[patient_col]) else f"환자_{index+1}"
            hospital_id  = str(row[hospital_col]).strip() if pd.notna(row[hospital_col]) else ''
            hospital_id  = ''.join(filter(str.isdigit, hospital_id))

            if len(hospital_id) != 8:
                progress_state['results'].insert(0, {'id': hospital_id, 'name': patient_name, 'status': '❌ 등록번호 형식 오류'})
                progress_state['current'] += 1
                continue

            progress_state['current'] = index + 1

            success, msg, _ = automate_registration(hospital_id, patient_name=patient_name)
            progress_state['results'].insert(0, {'id': hospital_id, 'name': patient_name, 'status': msg})

            if _stop_flag:
                progress_state['results'].insert(0, {'id': '-', 'name': '-', 'status': '🚨 ESC 키에 의해 작업이 중단되었습니다.'})
                break

    except _StopAutomation:
        progress_state['results'].insert(0, {'id': '-', 'name': '-', 'status': '🚨 ESC 키에 의해 작업이 중단되었습니다.'})
    except Exception as e:
        progress_state['results'].insert(0, {'id': 'ERROR', 'name': '오류 발생', 'status': str(e)})
    finally:
        progress_state['is_running'] = False
        try: os.remove(filepath)
        except: pass


@app.route('/register-excel', methods=['POST'])
def register_excel():
    global progress_state, _amis_hwnd_cache, _vsp_resized, _vsp_ctrl_cache
    try:
        if 'excel_file' not in request.files: return jsonify({'success': False, 'message': '파일이 없습니다.'})
        file = request.files['excel_file']
        if file.filename == '' or not allowed_file(file.filename): return jsonify({'success': False, 'message': '올바른 엑셀 파일이 아닙니다.'})

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        df = pd.read_excel(filepath, header=0)

        _amis_hwnd_cache = None
        _vsp_resized = False
        reset_stop_flag()
        _vsp_ctrl_cache = None

        progress_state = {
            'is_running': True,
            'total': len(df),
            'current': 0,
            'results': []
        }

        thread = threading.Thread(target=process_excel_background, args=(df, filepath))
        thread.start()

        return jsonify({'success': True, 'message': '작업 시작됨'})

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/progress', methods=['GET'])
def get_progress():
    return jsonify(progress_state)


def process_register_background(hospital_ids, patient_names):
    global progress_state
    try:
        for i, h_id in enumerate(hospital_ids):
            if _stop_flag:
                progress_state['results'].insert(0, {'id': '-', 'name': '-', 'status': '🚨 ESC 키에 의해 작업이 중단되었습니다.'})
                break
                
            p_name = patient_names[i] if i < len(patient_names) else "조회 중"
            progress_state['current'] = i + 1
            
            success, message, retrieved_name = automate_registration(h_id, patient_name=p_name)
            progress_state['results'].insert(0, {'id': h_id, 'name': retrieved_name, 'status': message})
            time.sleep(0.1)

    except _StopAutomation:
        progress_state['results'].insert(0, {'id': '-', 'name': '-', 'status': '🚨 ESC 키에 의해 작업이 중단되었습니다.'})
    except Exception as e:
        progress_state['results'].insert(0, {'id': 'ERROR', 'name': '오류 발생', 'status': str(e)})
    finally:
        progress_state['is_running'] = False


@app.route('/register', methods=['POST'])
def register():
    global progress_state, _amis_hwnd_cache, _vsp_resized, _vsp_ctrl_cache
    try:
        hospital_ids  = request.form.getlist('hospital_ids[]')
        patient_names = request.form.getlist('patient_names[]')

        if not hospital_ids:
            return jsonify({'success': False, 'message': '등록번호가 없습니다.'})

        _amis_hwnd_cache = None
        _vsp_resized = False
        reset_stop_flag()
        _vsp_ctrl_cache = None
        
        progress_state = {
            'is_running': True,
            'total': len(hospital_ids),
            'current': 0,
            'results': []
        }

        thread = threading.Thread(target=process_register_background, args=(hospital_ids, patient_names))
        thread.start()

        return jsonify({'success': True, 'message': '작업 시작됨'})

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/extract-info', methods=['POST'])
def extract_info():
    try:
        if 'text_data' in request.form:
            text = request.form.get('text_data', '')
            ids = extract_hospital_id(text)
            return jsonify({'success': True, 'extracted_ids': ids})

        if 'image_file' in request.files:
            file = request.files['image_file']
            img = Image.open(file.stream).convert('RGB')
            processed_img = preprocess_for_ocr(img)
            ocr_text = pytesseract.image_to_string(processed_img, lang='kor+eng')
            ids = extract_hospital_id(ocr_text)
            return jsonify({'success': True, 'extracted_ids': ids})

        return jsonify({'success': False, 'message': '전달된 데이터가 없습니다.'})

    except Exception as e:
        logger.error(f"추출 오류: {str(e)}")
        return jsonify({'success': False, 'message': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
