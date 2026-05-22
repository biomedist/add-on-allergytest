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

# Tesseract 설치 경로 지정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB 최대 파일 크기

# 업로드 폴더 생성
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# PyAutoGUI 설정
pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = True

# 로깅 설정
# 로그를 터미널 + 파일 동시 저장
# TimedRotatingFileHandler: 매월 1일 자동으로 새 파일 생성, 최대 12개월 보관
from logging.handlers import TimedRotatingFileHandler

_log_formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')

_stream_handler = logging.StreamHandler()
_stream_handler.setFormatter(_log_formatter)

_file_handler = TimedRotatingFileHandler(
    'ocr_log.txt',
    when='midnight',       # 자정마다 체크
    interval=30,           # 30일(약 한 달)마다 교체
    backupCount=12,        # 최대 12개 파일 보관 (약 1년치)
    encoding='utf-8'
)
_file_handler.setFormatter(_log_formatter)
_file_handler.suffix = '%Y%m'   # 파일명: ocr_log.txt.202604 형식

# basicConfig 대신 root logger에 직접 추가 (모든 스레드에서 동작)
_root_logger = logging.getLogger()
_root_logger.setLevel(logging.INFO)
_root_logger.addHandler(_stream_handler)
_root_logger.addHandler(_file_handler)

logger = logging.getLogger(__name__)

# 희망일자 체크 시 적용할 '분류' 값
ALLOWED_CLASSIFICATION = frozenset({'D52', 'D53', 'D54'})

# ── 긴급 정지 플래그 ──
_stop_flag = False

# [새로 추가] 진행 상황을 저장할 전역 딕셔너리
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
        raise InterruptedError("사용자가 ESC로 자동화를 중단했습니다.")

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

# ── uiautomation 헬퍼 함수들 ──

def _get_vsp_window_cached():
    """채혈접수 창 컨트롤 반환"""
    try:
        comtypes.CoInitialize()
        amis = auto.WindowControl(Name='AMIS3', timeout=1)
        vsp = auto.WindowControl(
            searchFromControl=amis,
            Name='채혈접수[VSPSSSP010S]_서관2층검사실',
            timeout=1
        )
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
        # 병원등록번호 입력창: 값이 숫자 8자리인 EditControl 탐색
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
    """uiautomation으로 환자명 읽기 - 등록번호가 있는 aPanelControl1 집중 탐색"""
    try:
        comtypes.CoInitialize()
        vsp = _get_vsp_window_cached()
        if not vsp:
            return ''

        korean_pattern = re.compile(r'^[가-힣]{2,5}$')

        # 로그에서 확인된 등록번호/환자정보 패널(aPanelControl1)로 직행
        panel = auto.PaneControl(searchFromControl=vsp, Name='aPanelControl1', timeout=1)
        
        # 패널 내부의 자식 컨트롤들만 빠르게 탐색
        for ctrl, depth in auto.WalkControl(panel, maxDepth=3):
            if ctrl.ControlTypeName in ['EditControl', 'TextControl', 'CustomControl']:
                # 1. Value 읽기 시도
                try:
                    val = ctrl.GetValuePattern().Value or ''
                    val = val.strip()
                    if korean_pattern.match(val):
                        logger.info(f"uiautomation 환자명 발견(Value): {val}")
                        return val
                except:
                    pass

                # 2. Name 읽기 시도
                name = ctrl.Name or ''
                name = name.strip()
                if korean_pattern.match(name):
                    logger.info(f"uiautomation 환자명 발견(Name): {name}")
                    return name

        return ''
    except Exception as e:
        logger.debug(f"uiautomation 환자명 읽기 실패 (클립보드로 넘어갑니다): {e}")
        return ''

def ui_click_action_and_register():
    """uiautomation으로 Action 버튼 클릭 후 채혈접수 클릭"""
    try:
        comtypes.CoInitialize()
        vsp = _get_vsp_window_cached()
        if not vsp:
            return False

        # Action 버튼 찾기 - 채혈접수 창(aPanelControl2)에서만 탐색
        # 탐색 결과: aPanelControl2 바로 아래에 ButtonControl 이름:'Action' 있음
        action_btn = None
        try:
            # scrollPanel > screenMainPanel 안의 aPanelControl2에서 탐색
            scroll_panel = auto.PaneControl(searchFromControl=vsp, Name='scrollPanel', timeout=1)
            panel2 = auto.PaneControl(searchFromControl=scroll_panel, Name='aPanelControl2', timeout=1)
            action_btn = auto.ButtonControl(searchFromControl=panel2, Name='Action', timeout=1)
            logger.info("aPanelControl2에서 Action 버튼 발견")
        except Exception as e:
            logger.warning(f"aPanelControl2 Action 탐색 실패: {e}, 전체 탐색 시도")
            # fallback: vsp 전체에서 탐색하되 채혈접수 창 기준
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

        logger.info("Action 버튼 클릭 (가려진 버튼 Invoke 시도)")
        try:
            # [핵심] 마우스 이동 없이 프로그램 내부적으로 '클릭 신호'를 보냅니다.
            action_btn.GetInvokePattern().Invoke()
        except Exception as e:
            logger.warning(f"Invoke 실패, 일반 클릭 시도: {e}")
            action_btn.Click() # 실패 시 물리적 클릭으로 폴백
            
        time.sleep(0.3)  # 드롭다운 열릴 때까지 대기

        # 채혈접수 = F4 단축키
        import pyautogui as _pag
        _pag.press('f4')
        logger.info("채혈접수 F4 키 입력 완료")
        return True

    except Exception as e:
        logger.warning(f"uiautomation Action 클릭 실패 (좌표 방식으로 fallback): {e}")
        return False


# ── AMIS 창 핸들 캐시 (첫 번째 환자 때 한 번만 찾고, 이후엔 재사용) ──
_amis_hwnd_cache = None

# ── uiautomation vsp 창 캐시 (매 환자마다 탐색하지 않도록) ──
_vsp_ctrl_cache = None

def _get_vsp_window_cached():
    """채혈접수 창을 캐시로 반환 (첫 번째만 탐색, 이후 재사용)"""
    global _vsp_ctrl_cache
    comtypes.CoInitialize()
    if _vsp_ctrl_cache is not None:
        try:
            _ = _vsp_ctrl_cache.Name  # 접근 가능하면 유효
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

        # 1단계: 최소화된 창이면 복원
        if win32gui.IsIconic(hwnd):
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) # 창이 최소화 되어 있다면 복구
                win32gui.SetForegroundWindow(hwnd) # 창을 맨 앞으로
            except Exception as e:
                print(f"오류: {e}")

            time.sleep(0.2)  # 최소화 복원 대기
        check_stop()

        # 2단계: 스레드 연결 (Windows 보안 우회)
        current_hwnd = win32gui.GetForegroundWindow()
        current_thread_id = win32api.GetCurrentThreadId()
        target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
        foreground_thread_id, _ = win32process.GetWindowThreadProcessId(current_hwnd)

        ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, True)
        ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, True)

        # 3단계: 창 활성화 및 크기 조정
        try: 
            import win32con 
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                logger.warning(f"SetForegroundWindow 실패: {e}")
                
            ctypes.windll.user32.SetFocus(hwnd)
            ctypes.windll.user32.SetActiveWindow(hwnd)

            # 4단계: 작업표시줄 제외 영역의 '좌측 2/3'로 크기 조정
            work_area = RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(work_area), 0)
            
            total_width = work_area.right - work_area.left
            # [수정] 너비를 전체의 2/3로 설정합니다. (Action 버튼 노출 충분함)
            target_width = (total_width * 2) // 3
            height = work_area.bottom - work_area.top
            
            try:
                win32gui.MoveWindow(hwnd, work_area.left, work_area.top, target_width, height, True)
                logger.info(f"창 크기 조정 완료: 좌측 2/3 배치 ({target_width}px)")
            except Exception as e:
                logger.warning(f"창 이동 실패: {e}")

            time.sleep(0.05)

            # 스레드 연결 해제 (정상 흐름)
            ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, False)
            ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, False)

            time.sleep(0.1)
            active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            logger.info(f"활성화된 창 제목: {active_title}")

            return True

        except Exception as e:
            logger.error(f"창 활성화 프로세스 중 오류(무시): {e}")
            # 에러 발생 시에도 스레드 연결 해제 시도
            try:
                ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, False)
                ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, False)
            except:
                pass
            
            # fallback: pygetwindow
            # 메인 Win32 방식 실패 시 pygetwindow로 마지막 시도
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
            
            return True # 모든 시도가 실패해도 자동화를 계속 진행하도록 True 반환

    except Exception as e:
        logger.error(f"find_window_by_title 전체 오류: {e}")
        return False


def find_amis_window():
    """AMIS 창을 찾아 활성화. hwnd를 캐싱해 두 번째 환자부터는 포커스만 확인."""
    global _amis_hwnd_cache
    import win32gui
    import win32con

    # ── 캐시된 핸들이 있으면 포커스 확인만 하고 즉시 반환 ──
    if _amis_hwnd_cache is not None:
        try:
            # 창이 여전히 유효한지 확인
            if win32gui.IsWindow(_amis_hwnd_cache) and win32gui.IsWindowVisible(_amis_hwnd_cache):
                active_hwnd = win32gui.GetForegroundWindow()
                if active_hwnd != _amis_hwnd_cache:
                    # 최소화 복원
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

    # ── 첫 실행이거나 캐시 무효화 시 → 전체 탐색 ──
    # "AMIS"가 포함된 창을 먼저 찾고, 실패 시 나머지 검색어로 시도
    search_terms = ["AMIS", "VSPSSSP010S", "채혈접수", "AMIS 3.0"]
    for term in search_terms:
        try:
            if find_window_by_title(term):
                # 찾은 hwnd를 캐시에 저장
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


# ── AMIS 내부 채혈접수 창 크기 조정 (일괄접수 세션당 1회) ──
_vsp_resized = False

def reset_vsp_resize():
    """일괄접수 시작 시 호출 → 세션마다 창 크기 재조정"""
    global _vsp_resized
    _vsp_resized = False

def resize_vsp_window():
    """채혈접수[VSPSSSP010S] MDI 내부 창을 30행이 보이도록 크기 조정"""
    global _vsp_resized
    if _vsp_resized:
        return  # 이미 조정했으면 스킵

    try:
        comtypes.CoInitialize()
        amis = auto.WindowControl(Name='AMIS3', timeout=2)
        vsp  = auto.WindowControl(
            searchFromControl=amis,
            Name='채혈접수[VSPSSSP010S]_서관2층검사실',
            timeout=2
        )

        # uiautomation으로 내부 창의 현재 Bounding Rectangle 확인
        rect = vsp.BoundingRectangle
        cur_left   = rect.left
        cur_top    = rect.top
        cur_width  = rect.width()
        cur_height = rect.height()

        # 목표: 높이를 충분히 늘려 30행 표시, 너비는 +20px(Action 버튼 노출)
        # 행 1개 높이 약 22px 기준 → 30행 = 660px + 헤더/여백 180px ≈ 840px
        TARGET_HEIGHT = max(cur_height, 840)
        TARGET_WIDTH  = cur_width + 20  # Action 버튼 노출용

        # WindowPattern으로 이동/크기 조정
        try:
            wp = vsp.GetWindowPattern()
            if wp.CanResize:
                # TransformPattern이 있으면 사용 (MDI 자식 창에 적합)
                tp = vsp.GetTransformPattern()
                if tp.CanResize:
                    tp.Resize(TARGET_WIDTH, TARGET_HEIGHT)
                    logger.info(f"채혈접수 창 크기 조정: {cur_width}x{cur_height} → {TARGET_WIDTH}x{TARGET_HEIGHT}")
                if tp.CanMove:
                    tp.Move(cur_left, cur_top)  # 위치 유지
        except Exception as e:
            logger.warning(f"TransformPattern 크기 조정 실패: {e}")

        _vsp_resized = True
        time.sleep(0.2)  # 리사이즈 반영 대기

    except Exception as e:
        logger.warning(f"채혈접수 창 크기 조정 실패 (무시하고 계속): {e}")
        _vsp_resized = True  # 실패해도 재시도 안 함


# ── Win32 작업표시줄 제외 영역 구조체 (함수 밖 전역 선언) ──
class RECT(ctypes.Structure):
    _fields_ = [
        ("left",   ctypes.c_long),
        ("top",    ctypes.c_long),
        ("right",  ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


# ─────────────────────────────────────────────
# 유틸리티 함수들
# ─────────────────────────────────────────────

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


# ─────────────────────────────────────────────
# 창 찾기 / 활성화
# ─────────────────────────────────────────────

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

        # 1단계: 최소화된 창이면 복원
        if win32gui.IsIconic(hwnd):
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) # 창이 최소화 되어 있다면 복구
                win32gui.SetForegroundWindow(hwnd) # 창을 맨 앞으로
            except Exception as e:
                print(f"오류: {e}")

            time.sleep(0.2)  # 최소화 복원 대기
        check_stop()

        # 2단계: 스레드 연결 (Windows 보안 우회)
        current_hwnd = win32gui.GetForegroundWindow()
        current_thread_id = win32api.GetCurrentThreadId()
        target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
        foreground_thread_id, _ = win32process.GetWindowThreadProcessId(current_hwnd)

        ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, True)
        ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, True)

        # 3단계: 창 활성화 (SyntaxError 해결 및 권한 우회 적용)
        try: 
            import win32con 
            # 창 복구 시도
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                logger.warning(f"SetForegroundWindow 실패(권한문제): {e}")
                
            ctypes.windll.user32.SetFocus(hwnd)
            ctypes.windll.user32.SetActiveWindow(hwnd)

        
            # 4단계: 작업표시줄 제외한 영역의 '좌측 2/3'로 크기 조정
            work_area = RECT() #
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(work_area), 0) #
            
            total_width = work_area.right - work_area.left
            # 너비를 전체의 2/3로 설정
            target_width = (total_width * 2) // 3
            height = work_area.bottom - work_area.top #
            
            try:
                win32gui.MoveWindow(hwnd, work_area.left, work_area.top, target_width, height, True) #
                logger.info(f"창 분할 완료: 좌측 2/3 ({target_width}px)")
            except Exception as e:
                logger.warning(f"창 이동 실패: {e}")

            time.sleep(0.05)

            # 스레드 연결 해제 (정상 흐름)
            ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, False)
            ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, False)

            time.sleep(0.1)
            active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            logger.info(f"활성화된 창 제목: {active_title}")

            return True

        except Exception as e:
            logger.error(f"창 활성화 프로세스 중 오류(무시): {e}")
            # 에러 발생 시에도 스레드 연결 해제 시도
            try:
                ctypes.windll.user32.AttachThreadInput(foreground_thread_id, current_thread_id, False)
                ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, False)
            except:
                pass
            
            # fallback: pygetwindow
            # 메인 Win32 방식 실패 시 pygetwindow로 마지막 시도
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
            
            return True # 모든 시도가 실패해도 자동화를 계속 진행하도록 True 반환

    except Exception as e:
        logger.error(f"find_window_by_title 전체 오류: {e}")
        return False


def find_amis_window_unused_placeholder():
    pass  # 위쪽 find_amis_window로 대체됨




# ─────────────────────────────────────────────
# 날짜 체크박스 클릭
# ─────────────────────────────────────────────

def click_date_checkboxes_in_range():
    """희망일자(오늘±3일) + 분류(D52/D53/D54) 조건 충족 행을 uiautomation으로 클릭"""
    today = datetime.now().date()

    # AMIS 희망일자 형식이 무엇이든 매칭되도록 여러 포맷 동시 생성
    valid_dates = set()
    for delta in range(-3, 4):
        d = today + timedelta(days=delta)
        valid_dates.add(d.strftime('%Y%m%d'))    # 20260318
        valid_dates.add(d.strftime('%Y-%m-%d'))  # 2026-03-18
        valid_dates.add(d.strftime('%Y/%m/%d'))  # 2026/03/18
        valid_dates.add(d.strftime('%y-%m-%d'))  # 26-03-18
        valid_dates.add(d.strftime('%y/%m/%d'))  # 26/03/18
        valid_dates.add(d.strftime('%m/%d/%Y'))  # 03/18/2026

    logger.info(f"유효 날짜 (오늘 기준 샘플): {today.strftime('%Y%m%d')} / {today.strftime('%Y-%m-%d')}")

    clicked_count = 0
    clicked_specimens = set()

    # ⚡ 전역 탐색 타임아웃을 0.5초로 단축 (기본값 보통 3~10초)
    auto.SetGlobalSearchTimeout(0.3)

    try:
        t0 = time.perf_counter()
        comtypes.CoInitialize()

        # ── 창 탐색: 직접 경로로 접근 ──
        amis = auto.WindowControl(Name='AMIS3', timeout=0.5)
        logger.info(f"[⏱] AMIS 창: {time.perf_counter()-t0:.3f}s")

        vsp = auto.WindowControl(searchFromControl=amis,
                                 Name='채혈접수[VSPSSSP010S]_서관2층검사실', timeout=0.5)
        logger.info(f"[⏱] VSP 창: {time.perf_counter()-t0:.3f}s")

        scroll_panel = auto.PaneControl(searchFromControl=vsp, Name='scrollPanel', timeout=1)
        logger.info(f"[⏱] scrollPanel: {time.perf_counter()-t0:.3f}s")

        # ── 처방목록 테이블: 재귀 탐색 대신 알려진 경로로 직접 접근 ──
        # 경로: scrollPanel > (첫 번째 TableControl)
        # GetChildren()은 즉시 반환되므로 재귀보다 훨씬 빠름
        target_table = None
        try:
            children = scroll_panel.GetChildren()
            logger.info(f"[⏱] scrollPanel.GetChildren() ({len(children)}개): {time.perf_counter()-t0:.3f}s")
            for child in children:
                if child.ControlTypeName == 'TableControl':
                    target_table = child
                    break
                # PaneControl 한 단계만 더 내려가서 확인
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

        # fallback: 기존 재귀 탐색 (단, depth 제한을 5로 강화)
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

        logger.info(f"[⏱] 테이블 탐색 완료: {time.perf_counter()-t0:.3f}s")

        if not target_table:
            logger.error("처방목록 테이블을 찾을 수 없습니다.")
            return 0

        # ── 데이터 패널 직접 접근 ──
        try:
            data_panel = auto.CustomControl(searchFromControl=target_table, Name='데이터 패널', timeout=1)
        except Exception:
            # '데이터 패널' CustomControl이 없으면 처방 없음
            logger.info("⚡ 데이터 패널 없음 → 처방목록 없음, 즉시 반환")
            return 0

        logger.info(f"[⏱] 데이터 패널 접근: {time.perf_counter()-t0:.3f}s")

        # -데이터 패널 접근 완료 후, 루프 시작 전 스크롤 초기화 ---
        try:
            rect = data_panel.BoundingRectangle
            center_x = rect.left + (rect.width() // 2)
            center_y = rect.top + (rect.height() // 2)
            # 마우스를 테이블 중앙으로 이동 후 위로 스크롤
            pyautogui.moveTo(center_x, center_y, duration=0)
            pyautogui.scroll(30) # 넉넉하게 위로 스크롤
            time.sleep(0.3)
            logger.info("⚡ 처방목록 스크롤 초기화 (맨 위로 이동)")
        except Exception as e:
            logger.warning(f"스크롤 초기화 실패: {e}")
        # --------------------------------------------------------------------------

        # ── 스크롤하며 전체 행 탐색 ──
        # AMIS 처방목록은 가상화 리스트: 화면에 보이는 행만 GetChildren()에 나타남.
        # → 행을 읽고 마지막 행에 도달하면 스크롤을 내려 다음 배치를 불러온다.

        # ── 스크롤 초기화: 맨 위로 이동 ──
        try:
            rect = data_panel.BoundingRectangle
            cx = rect.left + (rect.width() // 2)
            cy = rect.top  + (rect.height() // 2)
            pyautogui.moveTo(cx, cy, duration=0)
            pyautogui.scroll(30)  # 넉넉하게 위로
            time.sleep(0.1)
            logger.info("⚡ 처방목록 스크롤 맨 위로 초기화")
        except Exception as e:
            logger.warning(f"스크롤 초기화 실패: {e}")

        # ── 스크롤하며 전체 행 탐색 (행이 없을 때까지 무한 반복) ──
        seen_row_names = set()
        empty_streak   = 0
        scroll_attempts = 0

        while True:
            rows = data_panel.GetChildren()
            logger.info(f"[⏱] 총 {len(rows)}개 행 발견 (스크롤{scroll_attempts}): {time.perf_counter()-t0:.3f}s")

            if not rows:
                logger.info("⚡ 처방목록 없음 → 즉시 다음 환자로")
                break

            new_rows_found = False  # 이번 스크롤에서 새 행이 있었는지

            for row in rows:
                check_stop()
                try:
                    row_name = row.Name

                    # 이미 처리한 행은 건너뜀
                    if row_name in seen_row_names:
                        continue
                    seen_row_names.add(row_name)
                    new_rows_found = True

                    items = row.GetChildren()
                    date_val      = None
                    class_val     = None
                    has_any_value = False  # 이 행에서 값이 있는 컬럼이 하나라도 있는지
                    specimen_item = None
                    specimen_no   = None
                    all_col_log   = []     # 디버그: 모든 컬럼명+값 기록

                    for item in items:
                        item_name  = item.Name or ''
                        item_value = ''
                        try:
                            item_value = item.GetValuePattern().Value or ''
                        except:
                            pass
                        v = item_value.strip()

                        # 디버그용: 값이 있는 컬럼만 기록
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

                    # ── 빈 행 판단: 어떤 컬럼에도 값이 전혀 없을 때만 빈 행 ──
                    # (검사명 컬럼명이 AMIS 버전마다 달라 특정 컬럼 기준은 위험)
                    if not has_any_value:
                        empty_streak += 1
                        logger.info(f"⚡ {row_name} 완전 빈 행 ({empty_streak}번째)")
                        if empty_streak >= 1:
                            logger.info("⚡ 빈 행 발견 → 처방 데이터 끝, 즉시 탐색 종료")
                            new_rows_found = False  # while 루프도 종료
                            break
                        continue
                    else:
                        empty_streak = 0

                    # 첫 번째 스크롤에서 컬럼 구조 파악용 로그 (최초 5행만)
                    if scroll_attempts == 0 and len(seen_row_names) <= 5:
                        logger.info(f"📋 {row_name} 컬럼 구조: {' | '.join(all_col_log)}")

                    is_date_ok  = date_val  in valid_dates
                    is_class_ok = class_val in ALLOWED_CLASSIFICATION

                    logger.info(
                        f"🔍 {row_name} → 희망일자:[{date_val}] {'✅' if is_date_ok else '❌'}  |  "
                        f"분류:[{class_val}] {'✅' if is_class_ok else '❌'}"
                    )

                    if is_date_ok and is_class_ok:
                        if specimen_item is None:
                            # 검체번호 없음 = 아직 채혈 전 처방이거나 다른 검체 형태
                            # 행 전체를 클릭하면 체크 토글이 되어 해제될 수 있으므로 건너뜀
                            logger.info(f"   ⏭️ {row_name} 건너뜀 - 검체번호 없음 (날짜={date_val}, 분류={class_val})")
                            continue
                        if specimen_no in clicked_specimens:
                            logger.info(f"   ⏭️ {row_name} 건너뜀 - 중복 검체번호 ({specimen_no})")
                            continue

                        logger.info(f"🎯 조건 충족! {row_name} 클릭 (날짜={date_val}, 분류={class_val}, 검체번호={specimen_no})")
                        specimen_item.Click()
                        clicked_specimens.add(specimen_no)
                        clicked_count += 1
                        time.sleep(0.05)

                except Exception as e:
                    logger.debug(f"행 처리 중 오류: {e}")
                    continue

            # 이번 루프에서 새 행이 없었으면 마지막 행 도달 → 종료
            if not new_rows_found:
                logger.info("⚡ 새 행 없음 → 마지막 행 도달, 탐색 종료")
                break

            # 스크롤: 테이블 중앙에서 마우스 휠로 스크롤
            scroll_attempts += 1
            logger.info(f"마우스휠 스크롤 ({scroll_attempts}회)")
            try:
                rect = data_panel.BoundingRectangle
                center_x = rect.left + (rect.right - rect.left) // 2
                center_y = rect.top  + (rect.bottom - rect.top) // 2
                # 마우스를 테이블 위로 이동 후 휠 스크롤 (채혈접수 창 안에서만)
                pyautogui.moveTo(center_x, center_y, duration=0)
                pyautogui.scroll(-5)  # 음수 = 아래로 스크롤 (5클릭)
                time.sleep(0.1)
            except Exception as e:
                logger.warning(f"스크롤 실패: {e}")
                break

    except Exception as e:
        logger.error(f"uiautomation 테이블 읽기 실패: {e}")
        logger.error(traceback.format_exc())
        return 0
    finally:
        # 전역 타임아웃 복원
        auto.SetGlobalSearchTimeout(1)

    logger.info(f"✅ 클릭 완료: 총 {clicked_count}개 행")
    return clicked_count



# ─────────────────────────────────────────────
# 진단 도구 (좌표 확인용)
# ─────────────────────────────────────────────

def diagnose_coordinates():
    """
    AMIS 화면에서 날짜/분류 컬럼의 올바른 X 좌표를 찾는 진단 함수.
    사용법: Python 콘솔에서 from app import diagnose_coordinates; diagnose_coordinates()
    """
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


# ─────────────────────────────────────────────
# 메인 자동화 함수
# ─────────────────────────────────────────────

def automate_registration(hospital_id, patient_name=''):
    """등록번호를 받아서 자동화 작업을 수행하는 함수"""
    try:
        logger.info(f"자동화 시작: 등록번호 {hospital_id}")

        # 1단계: AMIS 창 활성화
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

        # 1.5단계: 채혈접수 내부 창 크기 조정 (최초 1회만 실행)
        resize_vsp_window()

        # 2~3단계: 등록번호 입력 (uiautomation 우선, 실패시 좌표 fallback)
        registration_input_x = 144
        registration_input_y = 89
        logger.info(f"2단계: 등록번호 입력: {hospital_id}")

        if not ui_input_hospital_id(hospital_id):
            # 좌표 방식 fallback
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

        # 4단계: Enter 후 처방목록 로딩 대기 (캐시된 vsp 사용, 최대 2초)
        pyautogui.press('enter')
        _wait_start = time.perf_counter()
        _vsp_for_wait = _get_vsp_window_cached()
        if _vsp_for_wait:
            for _ in range(20):  # 0.1초씩 최대 2초
                time.sleep(0.1)
                try:
                    _sp_chk = auto.PaneControl(
                        searchFromControl=_vsp_for_wait, Name='scrollPanel', timeout=0.2)
                    if _sp_chk.GetChildren():
                        logger.info(f"처방목록 로딩 완료 ({time.perf_counter()-_wait_start:.2f}s)")
                        break
                except:
                    pass
        else:
            time.sleep(0.5)  # vsp 캐시 없으면 고정 대기

        # 4.5단계: 환자명 읽기
        if not patient_name or patient_name == "이름 없음" or patient_name == "조회 중":
            logger.info("4.5단계: 환자명 읽기 시도")
            
            # 1차 시도: uiautomation (빠름)
            patient_name = ui_get_patient_name()
            
            # 2차 시도: 클립보드 백업 (로그에서 검증된 매우 정확한 좌표 사용)
            if not patient_name:
                try:
                    NAME_X, NAME_Y = 275, 90  # 검증된 환자명 좌표
                    pyautogui.doubleClick(NAME_X, NAME_Y)
                    time.sleep(0.05)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(0.05)
                    patient_name = pyperclip.paste().strip()
                    
                    # 한글 2~5글자인지 검증
                    if re.match(r'^[가-힣]{2,5}$', patient_name):
                        logger.info(f"클립보드 환자명 추출 성공: {patient_name}")
                    else:
                        patient_name = "확인불가"
                except Exception as e:
                    logger.warning(f"환자명 추출 최종 실패: {e}")
                    patient_name = "확인불가"


        check_stop()
        # 5단계: 희망일자 확인 및 체크박스 클릭
        logger.info("5단계: 희망일자 확인 및 체크박스 클릭")
        try:
            clicked_count = click_date_checkboxes_in_range()

            if clicked_count > 0:
                logger.info(f"{clicked_count}개의 조건을 충족하는 행을 클릭했습니다.")

                # 6~7단계: Action → 채혈접수 (uiautomation 우선, 실패시 좌표 fallback)
                if not ui_click_action_and_register():
                    pyautogui.click(2510, 144)
                    time.sleep(0.3)
                    pyautogui.click(2510, 167)

                return (True, "바코드를 출력하였습니다.", patient_name)

            else:
                logger.warning("충족되는 행이 없습니다.")
                # 입력란 클리어: pyautogui로 직접 처리 (컨트롤명 불일치 문제 회피)
                try:
                    pyautogui.click(registration_input_x, registration_input_y)
                    time.sleep(0.05)
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('delete')
                    logger.info("입력란 클리어 완료")
                except Exception as e:
                    logger.warning(f"입력란 클리어 실패(무시): {e}")
                return (True, "충족되는 행이 없습니다.", patient_name)

        except Exception as e:
            error_msg = f"날짜 체크박스 클릭 중 오류: {str(e)}"
            logger.error(error_msg)
            return (False, error_msg, patient_name)

    except InterruptedError as e:
        logger.warning(str(e))
        return (False, str(e), patient_name)
    except Exception as e:
        error_msg = f"자동화 프로세스 중 오류: {str(e)}"
        logger.error(error_msg)
        return (False, error_msg, patient_name)


# ─────────────────────────────────────────────
# Flask 라우트
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


# 백그라운드에서 엑셀을 한 줄씩 처리하는 함수
def process_excel_background(df, filepath):
    global progress_state
    try:
        # 환자명, 등록번호 컬럼 찾기 (기존 로직 동일)
        patient_col, hospital_col = df.columns[0], df.columns[1] if len(df.columns) > 1 else df.columns[0]
        for col in df.columns:
            col_lower = str(col).lower()
            if any(k in col_lower for k in ['환자', '이름', 'name', 'patient']): patient_col = col
            if any(k in col_lower for k in ['등록', '번호', 'id', 'hospital']): hospital_col = col

        for index, row in df.iterrows():
            patient_name = str(row[patient_col]).strip() if pd.notna(row[patient_col]) else f"환자_{index+1}"
            hospital_id  = str(row[hospital_col]).strip() if pd.notna(row[hospital_col]) else ''
            hospital_id  = ''.join(filter(str.isdigit, hospital_id))

            if len(hospital_id) != 8:
                progress_state['results'].insert(0, {'id': hospital_id, 'name': patient_name, 'status': '❌ 등록번호 형식 오류'})
                progress_state['current'] += 1
                continue

            # 현재 진행 순서 업데이트
            progress_state['current'] = index + 1

            # 자동화 함수 실행 (reset_stop_flag는 automate_registration 내부에서 처리)
            success, msg, _ = automate_registration(hospital_id, patient_name=patient_name)

            # 결과를 맨 앞에 추가 (최신 결과가 위에 보이도록)
            progress_state['results'].insert(0, {'id': hospital_id, 'name': patient_name, 'status': msg})

            # ESC 중단 체크는 한 환자 처리 완료 후에
            if _stop_flag:
                logger.warning("엑셀 처리 중 ESC 감지, 작업을 중단합니다.")
                progress_state['results'].insert(0, {'id': '-', 'name': '-', 'status': '🚨 ESC 키에 의해 작업이 중단되었습니다.'})
                break

    except Exception as e:
        progress_state['results'].insert(0, {'id': 'ERROR', 'name': '오류 발생', 'status': str(e)})
    finally:
        progress_state['is_running'] = False # 작업 종료 알림
        try: os.remove(filepath) # 임시 파일 삭제
        except: pass


@app.route('/register-excel', methods=['POST'])
def register_excel():
    """엑셀 파일 업로드 및 백그라운드 작업 시작"""
    global progress_state, _amis_hwnd_cache, _vsp_resized
    try:
        if 'excel_file' not in request.files: return jsonify({'success': False, 'message': '파일이 없습니다.'})
        file = request.files['excel_file']
        if file.filename == '' or not allowed_file(file.filename): return jsonify({'success': False, 'message': '올바른 엑셀 파일이 아닙니다.'})

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        df = pd.read_excel(filepath, header=0)

        # 세션 시작 시 캐시 및 리사이즈 플래그 초기화
        _amis_hwnd_cache = None
        _vsp_resized = False
        reset_stop_flag()
        global _vsp_ctrl_cache
        _vsp_ctrl_cache = None

        progress_state = {
            'is_running': True,
            'total': len(df),
            'current': 0,
            'results': []
        }

        # 작업을 백그라운드 스레드로 넘기고, 브라우저에는 즉시 성공 응답을 보냄
        import threading
        thread = threading.Thread(target=process_excel_background, args=(df, filepath))
        thread.start()

        return jsonify({'success': True, 'message': '작업 시작됨'})

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

# 프론트엔드에서 진행 상황을 물어볼 때 대답해주는 라우트
@app.route('/progress', methods=['GET'])
def get_progress():
    return jsonify(progress_state)

# 정보 추출 탭을 위한 백그라운드 작업 함수
def process_register_background(hospital_ids, patient_names):
    global progress_state
    try:
        for i, h_id in enumerate(hospital_ids):
            if _stop_flag:
                progress_state['results'].insert(0, {'id': '-', 'name': '-', 'status': '🚨 ESC 키에 의해 작업이 중단되었습니다.'})
                break
                
            p_name = patient_names[i] if i < len(patient_names) else "조회 중"
            progress_state['current'] = i + 1
            
            # 자동화 실행
            success, message, retrieved_name = automate_registration(h_id, patient_name=p_name)
            
            # 결과 업데이트
            progress_state['results'].insert(0, {'id': h_id, 'name': retrieved_name, 'status': message})
            time.sleep(0.1)
            
    except Exception as e:
        progress_state['results'].insert(0, {'id': 'ERROR', 'name': '오류 발생', 'status': str(e)})
    finally:
        progress_state['is_running'] = False


@app.route('/register', methods=['POST'])
def register():
    """정보추출 탭 일괄 처리 (비동기 백그라운드로 전환)"""
    global progress_state, _amis_hwnd_cache, _vsp_resized
    try:
        hospital_ids  = request.form.getlist('hospital_ids[]')
        patient_names = request.form.getlist('patient_names[]')

        if not hospital_ids:
            return jsonify({'success': False, 'message': '등록번호가 없습니다.'})

        # 세션 시작 시 캐시 및 리사이즈 플래그 초기화
        _amis_hwnd_cache = None
        _vsp_resized = False
        reset_stop_flag()
        global _vsp_ctrl_cache
        _vsp_ctrl_cache = None
        
        progress_state = {
            'is_running': True,
            'total': len(hospital_ids),
            'current': 0,
            'results': []
        }

        # 백그라운드 스레드 시작
        import threading
        thread = threading.Thread(target=process_register_background, args=(hospital_ids, patient_names))
        thread.start()

        # 즉시 응답을 보내 브라우저가 멈추지 않게 함
        return jsonify({'success': True, 'message': '작업 시작됨'})

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/extract-info', methods=['POST'])
def extract_info():
    """텍스트 또는 이미지에서 8자리 번호 추출"""
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