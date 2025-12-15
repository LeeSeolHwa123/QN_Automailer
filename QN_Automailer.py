import os, re, time, uuid, tempfile, traceback, datetime as dt
from datetime import datetime as _dt
import win32com.client as win32

# =========================
# 글로벌 설정값 (GUI 또는 설정 파일에서 주입)
# =========================
_GLOBAL_SETTINGS = {}
_PROGRESS_CALLBACK = None  # GUI의 진행률 콜백 함수
_MESSAGE_CALLBACK = None   # GUI의 상세 메시지 콜백 함수

def init_settings(settings: dict, progress_callback=None, message_callback=None):
    """
    글로벌 설정값 초기화 (GUI 또는 메인에서 호출)
    Args:
        settings: 설정 딕셔너리
        progress_callback: 진행 단계 콜백 함수 (step_index 받음)
        message_callback: 상세 메시지 콜백 함수 (step_index, message 받음)
    """
    global _GLOBAL_SETTINGS, _PROGRESS_CALLBACK, _MESSAGE_CALLBACK
    _GLOBAL_SETTINGS = settings or {}
    _PROGRESS_CALLBACK = progress_callback
    _MESSAGE_CALLBACK = message_callback

def report_progress(step_index: int, message: str = None):
    """진행 단계를 GUI에 보고"""
    print(f"[REPORT_PROGRESS] Step {step_index}")
    print(f"[REPORT_PROGRESS] Has MESSAGE_CALLBACK: {_MESSAGE_CALLBACK is not None}")
    
    if _PROGRESS_CALLBACK:
        try:
            _PROGRESS_CALLBACK(step_index)
            print(f"[REPORT_PROGRESS] Progress callback called")
        except Exception as e:
            print(f"[REPORT_PROGRESS_ERROR] Progress callback error: {e}")
    
    if message and _MESSAGE_CALLBACK:
        try:
            _MESSAGE_CALLBACK(step_index, message)
            print(f"[REPORT_PROGRESS] Message callback called")
        except Exception as e:
            print(f"[REPORT_PROGRESS_ERROR] Message callback error: {e}")
    elif message:
        print(f"[REPORT_PROGRESS] MESSAGE_CALLBACK is None")

# 기본값 (GUI 미사용 또는 폴백)
TEST_DATE = dt.date(2025, 12, 5)   # ← 원하는 테스트 날짜로 변경
USE_ACTIVE_MAIL_FIRST = True       # 열린 메일 우선 사용
MAIL_SUBJECT_KEY = "QN 등록 현황 공유"

CONFIG = {
    "EXPORT_DIR": r"C:\Users\seolhl\OneDrive - kochind.com\Documents\SAP\SAP GUI",
    "PLANT": "2201",
    # ZQCN - 첫 번째 표(C1-CK)
    "ZQCN_C_FIRST_TYPES": ["C1","C3","C5","C7","G1","G3","G5","G7","N1","CK"],
    # ZQCN - 두 번째 표(Q1~Q9)
    "ZQCN_Q_SECOND_TYPES": [f"Q{i}" for i in range(1,10)],
    # 파일명 규칙
    "FILE_FIRST": "{YYYYMMDD}.txt",
    "FILE_SECOND": "{YYYYMMDD}_Q1.txt",
    "FILE_ZNOT": "{YYYYMMDD}_ZNOT.txt",
}

# 글로벌 설정에서 현재값 추출
def get_test_date():
    """현재 실행 날짜 반환 (글로벌 설정 우선)."""
    if _GLOBAL_SETTINGS:
        try:
            from settings_manager import get_effective_execution_date
            return get_effective_execution_date(_GLOBAL_SETTINGS)
        except Exception:
            pass
    return dt.date.today()

def get_use_active_mail_first():
    """활성 메일 우선 사용 여부 (글로벌 설정 우선)"""
    if _GLOBAL_SETTINGS:
        return _GLOBAL_SETTINGS.get("use_active_mail_first", USE_ACTIVE_MAIL_FIRST)
    return USE_ACTIVE_MAIL_FIRST

def get_mail_subject_key():
    """메일 제목 키워드 (글로벌 설정 우선)"""
    if _GLOBAL_SETTINGS:
        return _GLOBAL_SETTINGS.get("mail_subject_key", MAIL_SUBJECT_KEY)
    return MAIL_SUBJECT_KEY

def get_config():
    """CONFIG 정보 (글로벌 설정 우선)"""
    if _GLOBAL_SETTINGS:
        global_config = CONFIG.copy()
        if _GLOBAL_SETTINGS.get("export_dir"):
            global_config["EXPORT_DIR"] = _GLOBAL_SETTINGS["export_dir"]
        if _GLOBAL_SETTINGS.get("plant"):
            global_config["PLANT"] = _GLOBAL_SETTINGS["plant"]
        return global_config
    return CONFIG

def get_date_range():
    """날짜 조회 범위 반환: (start_date, end_date)"""
    if _GLOBAL_SETTINGS:
        from settings_manager import get_effective_date_range
        return get_effective_date_range(_GLOBAL_SETTINGS)
    
    today = get_test_date()
    return dt.date(today.year, 1, 1), today

def get_mail_recipients():
    """메일 수신자 리스트 (글로벌 설정 우선)"""
    if _GLOBAL_SETTINGS and _GLOBAL_SETTINGS.get("mail_recipients"):
        from settings_manager import get_mail_recipients_list
        return get_mail_recipients_list(_GLOBAL_SETTINGS)
    return ["Byungjin.Hong@molex.com", "JoonYong.Son@molex.com", 
            "YongHyuk.Ahn@molex.com", "myeongjun.kim@molex.com"]


def apply_settings_to_pipeline(settings: dict):
    """설정값을 qn_full_pipeline 모듈에 주입"""
    try:
        import qn_full_pipeline as fp
    except ImportError:
        log("[WARN] qn_full_pipeline 모듈을 불러오지 못했습니다.")
        return None

    if not settings:
        return fp

    # 실행 날짜 / 메일 설정
    try:
        from settings_manager import get_effective_execution_date, get_effective_date_range, get_mail_recipients_list
        fp.TEST_DATE = get_effective_execution_date(settings)
        start_d, end_d = get_effective_date_range(settings)
        fp.START_DATE = start_d
        fp.END_DATE = end_d
        fp.MAIL_RECIPIENTS = get_mail_recipients_list(settings)
    except Exception:
        fp.TEST_DATE = dt.date.today()
        fp.START_DATE = None
        fp.END_DATE = None
        fp.MAIL_RECIPIENTS = fp.MAIL_RECIPIENTS if hasattr(fp, "MAIL_RECIPIENTS") else []

    fp.USE_ACTIVE_MAIL_FIRST = settings.get("use_active_mail_first", fp.USE_ACTIVE_MAIL_FIRST)
    fp.MAIL_SUBJECT_KEY = settings.get("mail_subject_key", fp.MAIL_SUBJECT_KEY)

    # Config 덮어쓰기
    cfg = fp.CONFIG.copy()
    if settings.get("export_dir"):
        cfg["EXPORT_DIR"] = settings["export_dir"]
    if settings.get("plant"):
        cfg["PLANT"] = settings["plant"]

    templates = settings.get("file_templates", {}) or {}
    cfg["FILE_FIRST"] = templates.get("first", cfg.get("FILE_FIRST", cfg["FILE_FIRST"]))
    cfg["FILE_SECOND"] = templates.get("second", cfg.get("FILE_SECOND", cfg["FILE_SECOND"]))
    cfg["FILE_ZNOT"] = templates.get("znot", cfg.get("FILE_ZNOT", cfg["FILE_ZNOT"]))

    if settings.get("types_first"):
        cfg["ZQCN_C_FIRST_TYPES"] = settings["types_first"]
    if settings.get("types_second"):
        cfg["ZQCN_Q_SECOND_TYPES"] = settings["types_second"]

    fp.CONFIG.update(cfg)

    # 주입 완료 플래그 설정 (pipeline이 settings_manager를 다시 읽지 않도록)
    try:
        fp._SETTINGS_APPLIED = True
    except Exception:
        pass

    log(
        f"[SETTINGS→PIPELINE] TEST_DATE={fp.TEST_DATE}, RANGE=({fp.START_DATE}~{fp.END_DATE}), "
        f"EXPORT_DIR={cfg.get('EXPORT_DIR')}, RECIPIENTS={fp.MAIL_RECIPIENTS}"
    )
    return fp

# =========================
# QN TYPE 집합 정의
# =========================
VALID_TYPES_FIRST = {"C1","C3","C5","C7","G1","G3","G5","G7","N1","CK"}
VALID_TYPES_SECOND = {f"Q{i}" for i in range(1,10)}

def simplify_qn_type(raw: str) -> str:
    if not raw:
        return ""
    raw = raw.strip().upper()

    if raw in VALID_TYPES_FIRST:
        return raw
    if raw in VALID_TYPES_SECOND:
        return raw

    m = re.match(r"^(Q[1-9]|C[1357]|G[1357]|N1|CK)", raw)
    if m:
        return m.group(1)

    return raw

    
def dstr(d: dt.date): return d.strftime("%Y%m%d")
def yday(d: dt.date): return d - dt.timedelta(days=1)

def path_first(day: dt.date):
    cfg = get_config()
    return os.path.join(cfg["EXPORT_DIR"], cfg["FILE_FIRST"].format(YYYYMMDD=dstr(day)))
def path_second(day: dt.date):
    cfg = get_config()
    return os.path.join(cfg["EXPORT_DIR"], cfg["FILE_SECOND"].format(YYYYMMDD=dstr(day)))
def path_znot(day: dt.date):
    cfg = get_config()
    return os.path.join(cfg["EXPORT_DIR"], cfg["FILE_ZNOT"].format(YYYYMMDD=dstr(day)))

def log(msg): print(msg, flush=True)
def log_exc(ctx=""): print(f"[EXC] {ctx}\n{traceback.format_exc()}", flush=True)

# =========================
# SAP COMMON
# =========================
def get_session():
    SapGuiAuto = win32.GetObject("SAPGUI")
    app = SapGuiAuto.GetScriptingEngine
    for ci in range(app.Children.Count):
        con = app.Children(ci)
        if con.DisabledByServer: continue
        for si in range(con.Children.Count):
            ses = con.Children(si)
            try:
                _ = ses.findById("wnd[0]")
                log(f"[SAP] 세션 연결됨: {ses.Info.SystemName}/{ses.Info.Client}/{ses.Info.User}")
                return ses
            except: pass
    raise RuntimeError("No active SAP session found.")

def wait_exists(session, obj_id, timeout=10, interval=0.2):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            return session.findById(obj_id)
        except:
            time.sleep(interval)
    return None

def safe_settext(session, obj_id, text):
    ctl = wait_exists(session, obj_id, timeout=8)
    if not ctl: raise RuntimeError(f"컨트롤 없음: {obj_id}")
    ctl.text = str(text)
    try: ctl.setFocus()
    except: pass

def safe_press(session, obj_id):
    btn = wait_exists(session, obj_id, timeout=8)
    if not btn: raise RuntimeError(f"버튼 없음: {obj_id}")
    btn.press()

def send_enter(session, wnd="wnd[0]"): session.findById(wnd).sendVKey(0)

def wait_idle(session, timeout=20):
    app = win32.GetObject("SAPGUI").GetScriptingEngine
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if not app.Busy and not session.Busy:
                return True
        except: pass
        time.sleep(0.15)
    return False

def close_any_popup(session):
    try:
        if session.FindById("wnd[1]", False):
            session.findById("wnd[1]").sendVKey(0)
            wait_idle(session, 5)
            log("[SAP] 팝업 Enter로 닫음")
            return True
    except: pass
    return False

def go_tcode(session, tcode):
    ok = wait_exists(session, "wnd[0]/tbar[0]/okcd", timeout=8)
    ok.text = f"/n{tcode}"
    send_enter(session)
    wait_idle(session, 10)
    log(f"[SAP] TCode 이동: {tcode}")

def set_zqcn_common_filters(session, start_date: dt.date, end_date: dt.date, plant: str):
    safe_settext(session, "wnd[0]/usr/ctxt[4]", start_date.strftime("%Y/%m/%d"))
    safe_settext(session, "wnd[0]/usr/ctxt[5]", end_date.strftime("%Y/%m/%d"))
    safe_settext(session, "wnd[0]/usr/ctxt[8]", plant)
    wait_exists(session, "wnd[0]/usr/chk[0]").selected = True
    wait_exists(session, "wnd[0]/usr/chk[1]").selected = True
    wait_exists(session, "wnd[0]/usr/chk[2]").selected = True
    wait_exists(session, "wnd[0]/usr/chk[3]").selected = True
    log(f"[ZQCN] 기간={start_date}~{end_date}, Plant={plant}")

def multisel_types(session, types):
    safe_press(session, "wnd[0]/usr/btn[1]")
    TABLE_BASE = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssub/1/2/tblSAPLALDBSINGLE"
    wait_exists(session, TABLE_BASE, timeout=10)

    first_page_count = min(8, len(types))

    for idx in range(first_page_count):
        cell_id = f"{TABLE_BASE}/ctxt[1,{idx}]"
        safe_settext(session, cell_id, types[idx])
        time.sleep(0.25)

    log(f"✓ 첫 페이지 입력 완료: {types[:first_page_count]}")

    if first_page_count > 0:
        try:
            last_cell = f"{TABLE_BASE}/ctxt[1,{first_page_count-1}]"
            session.findById(last_cell).setFocus()
            session.findById("wnd[1]").sendVKey(0)
        except Exception as e:
            log(f"⚠ 첫 페이지 커밋 실패: {e}")

    wait_idle(session, timeout=4)
    time.sleep(0.4)

    if len(types) > 8:
        try:
            log("▼ PageDown: 다음 페이지 이동")
            session.findById("wnd[1]").sendVKey(82)
            wait_idle(session, timeout=6)
            time.sleep(1.0)
        except Exception as e:
            log(f"⚠ PageDown 이동 실패: {e}")

        last_first_page_value = types[7]
        remaining = [last_first_page_value] + types[8:]

        log(f"→ 두 번째 페이지 입력: {remaining}")

        for idx, val in enumerate(remaining):
            try:
                cell_id = f"{TABLE_BASE}/ctxt[1,{idx}]"
                safe_settext(session, cell_id, val)
                time.sleep(0.3)
            except Exception as e:
                log(f"⚠ 두 번째 페이지 {val} 입력 실패: {e}")

        try:
            last_idx = len(remaining) - 1
            last_cell = f"{TABLE_BASE}/ctxt[1,{last_idx}]"
            session.findById(last_cell).setFocus()
            session.findById("wnd[1]").sendVKey(0)
        except Exception as e:
            log(f"⚠ 두 번째 페이지 커밋 실패: {e}")

        wait_idle(session, timeout=4)
        time.sleep(0.4)

    closed = False
    for _ in range(3):
        try:
            tbar = session.findById("wnd[1]/tbar[0]")
            for i in range(tbar.Children.Count):
                b = tbar.Children(i)
                tip = (getattr(b, "Tooltip", "") or "") + " " + (getattr(b, "Text", "") or "")
                if any(k in tip.lower() for k in ("ok", "확인", "apply", "적용")):
                    b.press()
                    time.sleep(0.2)
                    break
        except:
            pass

        try:
            if not session.FindById("wnd[1]", False):
                closed = True
                break
        except:
            closed = True
            break

        try:
            session.findById("wnd[1]").sendVKey(0)
        except:
            pass

    log(f"[ZQCN] 입력 완료 (팝업닫힘={closed})")

def execute_onli(session):
    close_any_popup(session)
    ok = wait_exists(session, "wnd[0]/tbar[0]/okcd", timeout=6)
    ok.text = "=ONLI"
    send_enter(session)
    wait_idle(session, 12)
    log("[ZQCN] 실행(=ONLI) 완료")

def export_list_text_with_tabs(session, out_dir: str, file_basename: str):
    if not os.path.isdir(out_dir): raise RuntimeError(f"경로 없음: {out_dir}")
    full_path = os.path.join(out_dir, file_basename)
    log(f"[Export] 저장 예정: {full_path}")

    session.findById("wnd[0]/tbar[0]/okcd").text = "%PC"
    send_enter(session)

    try:
        session.findById("wnd[1]/usr/ctxt[0]").text = out_dir
        session.findById("wnd[1]/usr/ctxt[1]").text = file_basename
    except:
        pass

    try: safe_press(session, "wnd[1]/tbar[0]/btn[0]")
    except: pass

    log("✅ 내보내기 완료")
    return full_path

def run_zqcn_and_export(session, start_date: dt.date, end_date: dt.date, notif_types, outfile_fullpath):
    go_tcode(session, "ZQCN")
    cfg = get_config()
    set_zqcn_common_filters(session, start_date, end_date, cfg["PLANT"])
    multisel_types(session, notif_types)
    execute_onli(session)
    export_list_text_with_tabs(session, CONFIG["EXPORT_DIR"], os.path.basename(outfile_fullpath))

def ensure_files_for_test_date(day: dt.date):
    """TEST_DATE 기준 3개 파일 없으면 SAP에서 실행해 생성."""
    ses = get_session()
    cfg = get_config()
    y1 = dt.date(day.year, 1, 1)

    f1 = path_first(day)
    if not os.path.isfile(f1):
        log("[RUN] ZQCN(C1~CK) 수행")
        run_zqcn_and_export(ses, y1, day, cfg["ZQCN_C_FIRST_TYPES"], f1)

    f2 = path_second(day)
    if not os.path.isfile(f2):
        log("[RUN] ZQCN(Q1~Q9) 수행")
        run_zqcn_and_export(ses, y1, day, cfg["ZQCN_Q_SECOND_TYPES"], f2)

    fz = path_znot(day)
    if not os.path.isfile(fz):
        log("[RUN] ZNOT(C1~CK) 수행")
        go_tcode(ses, "ZNOT")
        safe_settext(ses, "wnd[0]/usr/ctxt[0]", y1.strftime("%Y%m%d"))
        safe_settext(ses, "wnd[0]/usr/ctxt[1]", day.strftime("%Y%m%d"))
        export_list_text_with_tabs(ses, cfg["EXPORT_DIR"], os.path.basename(fz))

# =========================
# 파일 파싱 & 신규/삭제 판별
# =========================
QN_RE = re.compile(r"\b(\d{8})\b")
DATE_RE = re.compile(r"\b(\d{2}/\d{2}/\d{2})\b")

def parse_export_txt(filepath, anchor_qn):
    """SAP Export 텍스트를 파싱"""
    log(f"[PARSE] 파일 열기: {filepath}")
    if not os.path.isfile(filepath): raise RuntimeError(f"파일 없음: {filepath}")
    with open(filepath,"r",encoding="utf-8-sig",errors="ignore") as f:
        lines=[ln.rstrip("\r\n") for ln in f]
    log(f"[PARSE] 총 라인수: {len(lines)}")

    header_idx=None
    for i,ln in enumerate(lines):
        if "Notif" in ln and "Created" in ln:
            header_idx=i; break
    if header_idx is None:
        log("[PARSE] 헤더 미발견")
        return []

    rows=[]
    for ln in lines[header_idx+1:]:
        s=ln.strip()
        if not s or "Outstanding notification" in s: continue
        m=QN_RE.search(s); 
        if not m: continue
        qn=m.group(1)
        ds=DATE_RE.findall(s)
        if not ds: continue

        raw_date = ds[0]
        try:
            dt_obj = _dt.strptime(raw_date, "%y/%m/%d")
            created_fmt = dt_obj.strftime("%Y.%m.%d")
        except:
            created_fmt = raw_date

        rows.append({"QN":qn,"Created":created_fmt})
    
    log(f"[PARSE] 데이터 행 파싱={len(rows)}건")

    if not rows:
        return []

    if not anchor_qn:
        return rows

    seq=[r["QN"] for r in rows]
    try:
        idx = seq.index(str(anchor_qn).strip())
        new_rows = rows[idx+1:]
        log(f"[PARSE] 신규={len(new_rows)}건")
        return new_rows
    except ValueError:
        log("[PARSE] anchor 미일치")
        return rows

def load_qn_list(filepath):
    if not os.path.isfile(filepath): return []
    with open(filepath,"r",encoding="utf-8-sig",errors="ignore") as f:
        lines=[ln.rstrip("\r\n") for ln in f]
    got=[]
    for ln in lines:
        m = QN_RE.search(ln)
        if m: got.append(m.group(1))
    return got

def detect_deleted(prev_file, today_file):
    prev_q = set(load_qn_list(prev_file))
    today_q= set(load_qn_list(today_file))
    deleted = sorted(list(prev_q - today_q))
    log(f"[DELCHK] 삭제={deleted}")
    return deleted

# =========================
# Outlook 메일
# =========================
def get_active_mail():
    outlook = win32.Dispatch("Outlook.Application")
    insp = outlook.ActiveInspector()
    if insp is None: return None
    item = insp.CurrentItem
    return item if getattr(item,"Class",None) == 43 else None

def find_latest_qn_mail_before_date(target_date: dt.date):
    """테스트 날짜 '이전' 메일만 후보"""
    outlook = win32.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    sent = mapi.GetDefaultFolder(5)
    items = sent.Items
    items.Sort("[SentOn]", True)
    for i in range(1, min(200, items.Count)+1):
        it = items.Item(i)
        if getattr(it,"Class",None) != 43: continue
        subj = (getattr(it,"Subject","") or "")
        sent_on = getattr(it, "SentOn", None)
        sent_date = sent_on.date() if sent_on else None
        if MAIL_SUBJECT_KEY in subj and sent_date and (sent_date < target_date):
            log(f"[MAIL] 후보: {subj}")
            return it
    return None

# =========================
# 메인 파이프라인
# =========================
def main(settings: dict = None):
    """메인 함수"""
    if settings:
        global _PROGRESS_CALLBACK, _MESSAGE_CALLBACK
        existing_progress_cb = _PROGRESS_CALLBACK
        existing_message_cb = _MESSAGE_CALLBACK

        init_settings(settings)

        _PROGRESS_CALLBACK = existing_progress_cb
        _MESSAGE_CALLBACK = existing_message_cb

    print("[DEBUG] main() started")

    # qn_full_pipeline 설정 주입
    pipeline = apply_settings_to_pipeline(_GLOBAL_SETTINGS or settings)
    if pipeline is None:
        raise RuntimeError("qn_full_pipeline 모듈을 불러오지 못했습니다.")

    # GUI로 단계/메시지를 전달하도록 콜백 연결
    def _cb(step_idx, msg=None):
        report_progress(step_idx, msg)

    # 실제 전체 파이프라인 실행 (양쪽 콜백 모두 연결)
    pipeline.main(progress_cb=_cb, message_cb=_cb)

    log("✅ 전체 파이프라인 완료")


if __name__ == "__main__":
    print("[TEST] QN_Automailer 테스트")
    try:
        from settings_manager import load_settings, validate_settings
        print("✅ 설정 모듈 임포트 성공")
        
        test_settings = {
            'export_dir': r'C:\SAP\Export',
            'mail_recipients': 'test@company.com'
        }
        is_valid, msg = validate_settings(test_settings)
        print(f"✅ 설정 검증: {is_valid}")
        
        init_settings(test_settings)
        print(f"✅ 글로벌 설정 초기화 완료")
        
    except Exception as e:
        print(f"❌ 테스트 오류: {e}")
        import traceback
        traceback.print_exc()
