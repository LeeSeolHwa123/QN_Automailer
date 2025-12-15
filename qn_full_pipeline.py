import os, re, time, uuid, tempfile, traceback, datetime as dt
from datetime import datetime as _dt
from contextlib import contextmanager
import win32com.client as win32
import pywintypes

try:
    import pythoncom
except ImportError:  # pragma: no cover - pythoncom ships with pywin32 on Windows
    pythoncom = None

# =========================
# TEST CONFIG (← 이 날짜만 바꾸면 테스트 가능)
# =========================
DEFAULT_TEST_DATE = dt.date(2025, 12, 4)
TEST_DATE = DEFAULT_TEST_DATE      # ← 원하는 테스트 날짜로 변경 (주입 시 덮어씀)
DEFAULT_USE_ACTIVE_MAIL_FIRST = True
USE_ACTIVE_MAIL_FIRST = DEFAULT_USE_ACTIVE_MAIL_FIRST
DEFAULT_MAIL_SUBJECT_KEY = "QN 등록 현황 공유"
MAIL_SUBJECT_KEY = DEFAULT_MAIL_SUBJECT_KEY

# 사용자가 입력한 값 주입 시 활용되는 기본값들
START_DATE = None                   # 조회 시작일 (없으면 당해 1월 1일)
END_DATE = None                     # 조회 종료일 (없으면 TEST_DATE)
DEFAULT_MAIL_RECIPIENTS = [
    "Byungjin.Hong@molex.com",
    "JoonYong.Son@molex.com",
    "YongHyuk.Ahn@molex.com",
    "myeongjun.kim@molex.com",
]
MAIL_RECIPIENTS = list(DEFAULT_MAIL_RECIPIENTS)

# 설정 적용 여부 플래그 (QN_Automailer에서 이미 주입했다면 중복 적용하지 않음)
_SETTINGS_APPLIED = False

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
_DEFAULT_CONFIG = CONFIG.copy()
# =========================
# QN TYPE 집합 정의
# =========================
VALID_TYPES_FIRST = {"C1","C3","C5","C7","G1","G3","G5","G7","N1","CK"}
VALID_TYPES_SECOND = {f"Q{i}" for i in range(1,10)}

def simplify_qn_type(raw: str) -> str:
    if not raw:
        return ""
    raw = raw.strip().upper()

    # 정확히 일치하는 경우
    if raw in VALID_TYPES_FIRST:
        return raw
    if raw in VALID_TYPES_SECOND:
        return raw

    # 접두사 패턴 매칭 (Q1~, C3~, CK~, G5~, N1~ 등)
    m = re.match(r"^(Q[1-9]|C[1357]|G[1357]|N1|CK)", raw)
    if m:
        return m.group(1)

    # 그래도 없으면 원본 전체를 반환 (디버깅 목적)
    return raw

    
def dstr(d: dt.date): return d.strftime("%Y%m%d")
def yday(d: dt.date): return d - dt.timedelta(days=1)

def path_first(day: dt.date):
    return os.path.join(CONFIG["EXPORT_DIR"], CONFIG["FILE_FIRST"].format(YYYYMMDD=dstr(day)))
def path_second(day: dt.date):
    return os.path.join(CONFIG["EXPORT_DIR"], CONFIG["FILE_SECOND"].format(YYYYMMDD=dstr(day)))
def path_znot(day: dt.date):
    return os.path.join(CONFIG["EXPORT_DIR"], CONFIG["FILE_ZNOT"].format(YYYYMMDD=dstr(day)))

def log(msg): print(msg, flush=True)
def log_exc(ctx=""): print(f"[EXC] {ctx}\n{traceback.format_exc()}", flush=True)


@contextmanager
def sap_com_apartment():
    """Ensure pythoncom initializes on worker threads before using SAP COM APIs."""

    if pythoncom is None:
        log("[WARN] pythoncom 모듈을 찾지 못했습니다. SAP 스크립팅이 실패할 수 있습니다.")
        yield
        return

    try:
        pythoncom.CoInitialize()
    except Exception as exc:
        log(f"[WARN] pythoncom.CoInitialize 실패: {exc}")
        yield
        return

    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception as exc:
            log(f"[WARN] pythoncom.CoUninitialize 실패: {exc}")

# =========================
# SAP COMMON (세션/대기/입력/실행/내보내기)
# =========================
def get_session():
    """
    활성 SAP GUI 세션을 반환. SAPGUI ROT가 없거나 세션이 없을 때 상세 안내를 던진다.
    """
    def _fetch_app():
        # 1) 이미 열린 ROT 객체 시도
        try:
            return win32.GetObject("SAPGUI").GetScriptingEngine
        except Exception:
            pass
        # 2) Dispatch로 재시도 (일부 환경에서 ROT 등록이 늦을 수 있음)
        try:
            return win32.Dispatch("SAPGUI").GetScriptingEngine
        except Exception as e:
            raise RuntimeError(
                "SAP GUI COM 객체를 찾지 못했습니다. SAP Logon을 실행해 세션을 연 뒤, 같은 비트수(32/64) Python에서 다시 실행하세요."
            ) from e

    app = _fetch_app()

    # SAP GUI가 Busy일 수 있으니 잠시 대기하며 탐색
    for _ in range(3):
        try:
            for ci in range(app.Children.Count):
                con = app.Children(ci)
                if getattr(con, "DisabledByServer", False):
                    continue
                for si in range(con.Children.Count):
                    ses = con.Children(si)
                    try:
                        _ = ses.findById("wnd[0]")
                        log(f"[SAP] 세션 연결됨: {ses.Info.SystemName}/{ses.Info.Client}/{ses.Info.User} Tx={ses.Info.Transaction}")
                        return ses
                    except Exception:
                        continue
        except Exception:
            time.sleep(0.3)
            continue

    raise RuntimeError(
        "활성 SAP 세션을 찾지 못했습니다. SAP Logon에서 대상 시스템에 로그인해 메인 창(wnd[0])이 열린 상태인지 확인하고, 스크립팅을 활성화한 뒤 다시 실행하세요."
    )

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

def page_down(session, wnd="0", repeat=1):
    for _ in range(repeat):
        session.findById(f"wnd[{wnd}]").sendVKey(82)

def set_zqcn_common_filters(session, start_date: dt.date, end_date: dt.date, plant: str):
    safe_settext(session, "wnd[0]/usr/ctxt[4]", start_date.strftime("%Y/%m/%d"))
    safe_settext(session, "wnd[0]/usr/ctxt[5]", end_date.strftime("%Y/%m/%d"))
    safe_settext(session, "wnd[0]/usr/ctxt[8]", plant)
    # 상태 4개 모두 체크
    wait_exists(session, "wnd[0]/usr/chk[0]").selected = True  # Open
    wait_exists(session, "wnd[0]/usr/chk[1]").selected = True  # Postponed
    wait_exists(session, "wnd[0]/usr/chk[2]").selected = True  # Being processed
    wait_exists(session, "wnd[0]/usr/chk[3]").selected = True  # Completed
    log(f"[ZQCN] 기간={start_date}~{end_date}, Plant={plant} 상태=ALL")

def multisel_types(session, types):

    safe_press(session, "wnd[0]/usr/btn[1]")
    TABLE_BASE = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssub/1/2/tblSAPLALDBSINGLE"
    wait_exists(session, TABLE_BASE, timeout=10)

    # ============================================================
    # 1) 첫 페이지: 첫 8개 입력
    # ============================================================
    first_page_count = min(8, len(types))

    for idx in range(first_page_count):
        cell_id = f"{TABLE_BASE}/ctxt[1,{idx}]"
        safe_settext(session, cell_id, types[idx])
        time.sleep(0.25)

    log(f"✓ 첫 페이지 입력 완료: {types[:first_page_count]}")

    # 커밋 (Enter)
    if first_page_count > 0:
        try:
            last_cell = f"{TABLE_BASE}/ctxt[1,{first_page_count-1}]"
            session.findById(last_cell).setFocus()
            session.findById("wnd[1]").sendVKey(0)
        except Exception as e:
            log(f"⚠ 첫 페이지 커밋 실패: {e}")

    wait_idle(session, timeout=4)
    time.sleep(0.4)


    # ============================================================
    # 2) 남은 타입이 없으면 종료
    # ============================================================
    if len(types) <= 8:
        pass
    else:
        # ============================================================
        # 3) PageDown 후 두 번째 페이지에 입력
        # ============================================================
        try:
            log("▼ PageDown: 다음 페이지 이동")
            session.findById("wnd[1]").sendVKey(82)
            wait_idle(session, timeout=6)
            time.sleep(1.0)
        except Exception as e:
            log(f"⚠ PageDown 이동 실패: {e}")

        # 🔥 두 번째 페이지 입력값 구성
        last_first_page_value = types[7]   # 첫 페이지 마지막 값
        remaining = [last_first_page_value] + types[8:]

        log(f"→ 두 번째 페이지 입력: {remaining}")

        for idx, val in enumerate(remaining):
            try:
                cell_id = f"{TABLE_BASE}/ctxt[1,{idx}]"
                safe_settext(session, cell_id, val)
                time.sleep(0.3)
            except Exception as e:
                log(f"⚠ 두 번째 페이지 {val} 입력 실패: {e}")

        # 커밋
        try:
            last_idx = len(remaining) - 1
            last_cell = f"{TABLE_BASE}/ctxt[1,{last_idx}]"
            session.findById(last_cell).setFocus()
            session.findById("wnd[1]").sendVKey(0)
        except Exception as e:
            log(f"⚠ 두 번째 페이지 커밋 실패: {e}")

        wait_idle(session, timeout=4)
        time.sleep(0.4)


    # ============================================================
    # 4) 팝업 닫기
    # ============================================================
    closed = False
    for _ in range(3):
        try:
            tbar = session.findById("wnd[1]/tbar[0]")
            for i in range(tbar.Children.Count):
                b = tbar.Children(i)
                tip = (getattr(b, "Tooltip", "") or "") + " " + (getattr(b, "Text", "") or "")
                if any(k in tip.lower() for k in ("ok", "확인", "apply", "적용", "copy", "복사", "transfer", "전달")):
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
    log("[ZQCN] 실행(=ONLI) 완료 → 결과화면 도달 시 내보내기 진행")

def try_select_radio(session, ids):
    for rid in ids:
        try:
            session.findById(rid).select()
            return True
        except: pass
    return False

def export_list_text_with_tabs(session, out_dir: str, file_basename: str):
    if not os.path.isdir(out_dir): raise RuntimeError(f"경로 없음: {out_dir}")
    full_path = os.path.join(out_dir, file_basename)
    log(f"[Export] 저장 예정: {full_path}")

    session.findById("wnd[0]/tbar[0]/okcd").text = "%PC"
    send_enter(session)

    fmt = wait_exists(session, "wnd[1]", timeout=5)
    if fmt:
        radio_candidates = [
            "wnd[1]/usr/sub/2/sub/2/1/rad[1,0]",
            "wnd[1]/usr/rad[1,0]",
            "wnd[1]/usr/ssub/2/2/rad[1,0]",
        ]
        if try_select_radio(session, radio_candidates):
            try: safe_press(session, "wnd[1]/tbar[0]/btn[0]")  # Generate
            except: safe_press(session, "wnd[1]/tbar[0]/btn[2]")

    path_popup = wait_exists(session, "wnd[1]/usr", timeout=6)
    if not path_popup:
        log("⚠ 경로/파일명 팝업이 안 뜸 → 바로 저장 설정일 수 있음")
        return full_path

    try:
        session.findById("wnd[1]/usr/ctxt[0]").text = out_dir
        session.findById("wnd[1]/usr/ctxt[1]").text = file_basename
    except:
        usr = session.findById("wnd[1]/usr")
        for c in usr.Children:
            try:
                cid = c.Id.lower()
                if ("dy_path" in cid or "ctxt[0]" in cid) and hasattr(c, "text"):
                    c.text = out_dir
                if ("dy_filename" in cid or "ctxt[1]" in cid) and hasattr(c, "text"):
                    c.text = file_basename
            except: pass

    try: safe_press(session, "wnd[1]/tbar[0]/btn[0]")
    except: safe_press(session, "wnd[1]/tbar[0]/btn[2]")

    log("✅ 내보내기 완료")
    return full_path

def run_zqcn_and_export(session, start_date: dt.date, end_date: dt.date, notif_types, outfile_fullpath):
    go_tcode(session, "ZQCN")
    set_zqcn_common_filters(session, start_date, end_date, CONFIG["PLANT"])
    multisel_types(session, notif_types)
    execute_onli(session)
    export_list_text_with_tabs(session, CONFIG["EXPORT_DIR"], os.path.basename(outfile_fullpath))

def ensure_files_for_test_date(day: dt.date, start_date: dt.date | None = None):
    """지정된 날짜 범위로 3개 파일 없으면 SAP에서 실행해 생성."""
    ses = get_session()
    y1 = start_date or dt.date(day.year, 1, 1)

    f1 = path_first(day)
    if not os.path.isfile(f1):
        log("[RUN] ZQCN(C1~CK) 수행 → 첫 번째 표 파일 생성")
        run_zqcn_and_export(ses, y1, day, CONFIG["ZQCN_C_FIRST_TYPES"], f1)

    f2 = path_second(day)
    if not os.path.isfile(f2):
        log("[RUN] ZQCN(Q1~Q9) 수행 → 두 번째 표 파일 생성")
        run_zqcn_and_export(ses, y1, day, CONFIG["ZQCN_Q_SECOND_TYPES"], f2)

    fz = path_znot(day)
    if not os.path.isfile(fz):
        log("[RUN] ZNOT(C1~CK) 수행 → ZNOT 파일 생성")
        go_tcode(ses, "ZNOT")
        # 날짜
        safe_settext(ses, "wnd[0]/usr/ctxt[0]", y1.strftime("%Y%m%d"))
        safe_settext(ses, "wnd[0]/usr/ctxt[1]", day.strftime("%Y%m%d"))
        # Notification Type Multiple (C1~CK)
        safe_press(ses, "wnd[0]/usr/btn[1]")
        table = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssub/1/2/tblSAPLALDBSINGLE"
        wait_exists(ses, table, timeout=10)
        vals = CONFIG["ZQCN_C_FIRST_TYPES"]
        for idx in range(min(8, len(vals))):
            ses.findById(f"{table}/ctxt[1,{idx}]").Text = vals[idx]
        if len(vals) > 8:
            ses.findById(table).VerticalScrollbar.Position += 1
            for idx2, v in enumerate(vals[8:]):
                ses.findById(f"{table}/ctxt[1,{idx2}]").Text = v
        # Copy(F8)
        safe_press(ses, "wnd[1]/tbar[0]/btn[8]")
        # Plant
        safe_settext(ses, "wnd[0]/usr/ctxt[6]", CONFIG["PLANT"])
        # Execute(F8)
        safe_press(ses, "wnd[0]/tbar[1]/btn[8]")
        export_list_text_with_tabs(ses, CONFIG["EXPORT_DIR"], os.path.basename(fz))

# =========================
# 파일 파싱 & 신규/삭제 판별
# =========================
QN_RE = re.compile(r"\b(\d{8})\b")
DATE_RE = re.compile(r"\b(\d{2}/\d{2}/\d{2})\b")

def _split_export_line(line):
    parts = line.split("\t")
    return [p.replace("\xa0", " ").strip() for p in parts]


def _find_col_idx(header_cols, keywords):
    def norm(text):
        return re.sub(r"[^a-z0-9#]", "", text.lower())
    targets = [norm(k) for k in keywords]
    for idx, col in enumerate(header_cols):
        token = norm(col)
        if any(t in token for t in targets):
            return idx
    return None


def _read_export_rows(filepath):
    log(f"[PARSE] 파일 열기: {filepath}")
    if not os.path.isfile(filepath): raise RuntimeError(f"파일 없음: {filepath}")
    with open(filepath,"r",encoding="utf-8-sig",errors="ignore") as f:
        lines=[ln.rstrip("\r\n") for ln in f]
    log(f"[PARSE] 총 라인수: {len(lines)}")

    header_idx=None
    header_text=""
    for i,ln in enumerate(lines):
        if "Notif" in ln and "Created" in ln:
            header_idx=i; header_text=ln; break
    if header_idx is None:
        log("[PARSE] 헤더(Notif/Created) 미발견 → 데이터 0")
        return []

    header_cols = _split_export_line(header_text)
    qn_col = _find_col_idx(header_cols, ["notif", "notification", "qn", "qn#"])
    created_col = _find_col_idx(header_cols, ["created", "creation", "등록", "regdate"])

    log(
        f"[PARSE] 헤더 인덱스={header_idx} | QN컬럼={qn_col} | Created컬럼={created_col} | 헤더='{header_text}'"
    )
    rows=[]
    for raw_line in lines[header_idx+1:]:
        s=raw_line.strip()
        if not s or "Outstanding notification" in s:
            continue

        cols = _split_export_line(raw_line)
        qn_raw = ""
        if qn_col is not None and qn_col < len(cols):
            qn_raw = cols[qn_col]
        if not qn_raw:
            m = QN_RE.search(s)
            if m:
                qn_raw = m.group(1)
        qn = normalize_qn(qn_raw)
        if not qn:
            continue

        created_raw = ""
        if created_col is not None and created_col < len(cols):
            created_raw = cols[created_col]
        if not created_raw:
            ds=DATE_RE.findall(s)
            if ds:
                created_raw = ds[0]
        if not created_raw:
            continue

        try:
            dt_obj = _dt.strptime(created_raw, "%y/%m/%d")
            created_fmt = dt_obj.strftime("%Y.%m.%d")
        except Exception as e:
            try:
                dt_obj = _dt.strptime(created_raw, "%Y.%m.%d")
                created_fmt = dt_obj.strftime("%Y.%m.%d")
            except Exception:
                log(f"[PARSE] 날짜 변환 실패: {created_raw} ({e})")
                created_fmt = created_raw

        rows.append({"QN":qn,"Created":created_fmt})
    log(f"[PARSE] 데이터 행 파싱={len(rows)}건")

    return rows


def parse_export_txt(filepath, anchor_qn, anchor_candidates=None):
    """
    SAP Export 텍스트를 파싱해서
    - QN: 8자리
    - Created: 'YYYY.MM.DD' 형식으로 변환하여 반환
    """
    rows = _read_export_rows(filepath)
    log(f"[PARSE] anchor_qn={anchor_qn} candidates={anchor_candidates}")

    if not rows:
        return []

    candidate_norms = []
    primary = normalize_qn(anchor_qn)
    if primary:
        candidate_norms.append(primary)

    if anchor_candidates:
        for cand in anchor_candidates:
            c_norm = normalize_qn(cand)
            if c_norm and c_norm not in candidate_norms:
                candidate_norms.append(c_norm)

    if not candidate_norms:
        log(f"[PARSE] anchor 없음 → 신규={len(rows)}건 (예: {rows[:3]})")
        return rows

    seq=[r["QN"] for r in rows]
    anchor_idx = None
    matched_anchor = None
    for cand in candidate_norms:
        if cand in seq:
            anchor_idx = seq.index(cand)
            matched_anchor = cand
            break

    if anchor_idx is None:
        log(f"[PARSE] anchor 미일치 → 후보 {candidate_norms} 중 없음 → 전체 신규 처리")
        log(f"[PARSE] 신규={len(rows)}건 (예: {rows[:3]})")
        return rows

    new_rows = rows[anchor_idx+1:]
    log(
        f"[PARSE] anchor '{matched_anchor}' 위치={anchor_idx} → 신규={len(new_rows)}건 "
        f"(예: {new_rows[:3]})"
    )
    return new_rows

def load_qn_list(filepath):
    pairs = load_qn_pairs(filepath)
    return [qn for _, qn in pairs]


def load_qn_pairs(filepath):
    if not os.path.isfile(filepath):
        return []

    try:
        rows = _read_export_rows(filepath)
    except Exception as exc:
        log(f"[LOAD_QN] '{filepath}' 파싱 실패: {exc}")
        return []

    pairs = []
    for r in rows:
        qn = normalize_qn(r.get("QN"))
        created = r.get("Created", "")
        if qn:
            pairs.append((created, qn))
    log(f"[LOAD_QN] {os.path.basename(filepath)} → QN쌍 {len(pairs)}건")
    return pairs

def detect_deleted(prev_file, today_file):
    prev_pairs = set(load_qn_pairs(prev_file))
    today_pairs = set(load_qn_pairs(today_file))

    deleted_pairs = sorted(list(prev_pairs - today_pairs), key=lambda x: x[1])
    deleted_qns = [qn for _, qn in deleted_pairs]

    log(
        f"[DELCHK] prev={os.path.basename(prev_file)}({len(prev_pairs)}) → "
        f"today={os.path.basename(today_file)}({len(today_pairs)}) / 삭제추정={deleted_pairs}"
    )
    return deleted_qns

# =========================
# Outlook (메일 소스/복제/날짜 치환)
# =========================
def _get_outlook_namespace(max_attempts=5, delay=0.5):
    """Outlook MAPI 네임스페이스를 안전하게 획득."""
    last_exc = None
    outlook = win32.Dispatch("Outlook.Application")
    for attempt in range(1, max_attempts + 1):
        try:
            mapi = outlook.GetNamespace("MAPI")
            # 간단한 헬스 체크: 받은 편지함 접근
            mapi.GetDefaultFolder(6)
            return outlook, mapi
        except pywintypes.com_error as exc:
            last_exc = exc
            hresult = exc.args[0] if exc.args else None
            if hresult == -2147023170:
                time.sleep(delay * attempt)
                outlook = win32.Dispatch("Outlook.Application")
                continue
            raise

    raise RuntimeError("Outlook MAPI 초기화 실패: 원격 프로시저를 호출하지 못했습니다.") from last_exc


def get_active_mail():
    outlook = win32.Dispatch("Outlook.Application")
    insp = outlook.ActiveInspector()
    if insp is None: return None
    item = insp.CurrentItem
    return item if getattr(item,"Class",None) == 43 else None

def find_latest_qn_mail_before_date(target_date: dt.date):
    """테스트 날짜 '이전' 메일만 후보"""
    outlook, mapi = _get_outlook_namespace()
    sent = mapi.GetDefaultFolder(5)  # Sent
    items = sent.Items
    items.Sort("[SentOn]", True)
    try: items.IncludeRecurrences=True
    except: pass
    for i in range(1, min(200, items.Count)+1):
        it = items.Item(i)
        if getattr(it,"Class",None) != 43: continue
        subj = (getattr(it,"Subject","") or "")
        sent_on = getattr(it, "SentOn", None)
        sent_date = sent_on.date() if sent_on else None
        if MAIL_SUBJECT_KEY in subj and sent_date and (sent_date < target_date):
            log(f"[MAIL] 후보: {subj} | SentOn={sent_on}")
            return it
    return None


def replace_date_in_subject(subject, src_date, dst_date):
    if not subject: return subject
    sm, sd = src_date.month, src_date.day
    dm, dd = dst_date.month, dst_date.day
    pat = rf"(?<!\d)0?{sm}/0?{sd}(?!\d)"
    return re.sub(pat, f"{dm}/{dd}", subject)

def replace_date_in_body(html, src_date, dst_date):
    if not html: return html
    sy, sm, sd = src_date.year, src_date.month, src_date.day
    dy, dm, dd = dst_date.year, dst_date.month, dst_date.day
    sep = r"(?:\s|&nbsp;|<[^>]+>)*"
    pat = rf"{sy}{sep}년{sep}{sm}{sep}월{sep}{sd}{sep}일"
    return re.sub(pat, f"{dy}년 {dm}월 {dd}일", html)

# =========================
# 메일 표 유틸 + 진단
# =========================
def get_tables(html): 
    t = re.findall(r"(?is)<table\b.*?>.*?</table>", html)
    log(f"[HTML] table count={len(t)}")
    return t

def table_rows(html_table): 
    rows = re.findall(r"(?is)<tr\b.*?>.*?</tr>", html_table)
    return rows

def strip_html(s): return re.sub(r"(?is)<.*?>","",s).replace("\xa0"," ").strip()
def header_cells(row_html): return [strip_html(x) for x in re.findall(r"(?is)<t[dh]\b.*?>.*?</t[dh]>", row_html)]
def data_cells(row_html): return [strip_html(x) for x in re.findall(r"(?is)<t[dh]\b.*?>.*?</t[dh]>", row_html)]
def rebuild_table(rows_html): return "<table>" + "".join(rows_html) + "</table>"

def find_qn_index_from_header(hdrs):
    for i, h in enumerate(hdrs):
        H = h.upper().replace(" ", "")
        if "QN" in H or "NOTIF" in H or "NOTIFICATION" in H or "QN#" in H:
            return i
    return None

def collect_qn_candidates_from_table(table_html, limit=20):
    rows = table_rows(table_html)
    log(f"[TBL] rows={len(rows)} (헤더 포함)")
    if len(rows) < 2:
        return []

    primary = header_cells(rows[0])
    secondary = header_cells(rows[1]) if len(rows) > 1 else []
    hdrs = _merge_header_rows(primary, secondary)
    qn_idx = find_qn_index_from_header(hdrs)
    log(f"[TBL] 헤더셀={len(hdrs)} qn_idx={qn_idx} hdrs={hdrs}")
    if qn_idx is None:
        return []

    candidates = []
    seen = set()
    for r in range(len(rows) - 1, 0, -1):
        cells = data_cells(rows[r])
        if not any(cells):
            continue
        if qn_idx >= len(cells):
            continue
        qn = normalize_qn(cells[qn_idx])
        if not qn or qn in seen:
            continue
        candidates.append(qn)
        seen.add(qn)
        log(f"[TBL] QN 후보 추가: idx={r} qn={qn}")
        if len(candidates) >= limit:
            break

    log(f"[TBL] 수집된 QN 후보 {len(candidates)}건")
    return candidates


def extract_last_qn_from_table(table_html):
    candidates = collect_qn_candidates_from_table(table_html, limit=1)
    return candidates[0] if candidates else None

def count_data_rows(table_html):
    rows = table_rows(table_html)
    return max(0, len(rows)-1)

def remove_deleted_from_table(table_html, deleted_qns):
    before_rows = count_data_rows(table_html)
    log(f"[MAIL/DEL] 삭제처리 시작 (before_rows={before_rows}, deleted_qns={deleted_qns})")
    if not deleted_qns:
        return table_html, []

    rows = table_rows(table_html)
    if len(rows) < 2:
        return table_html, []

    header_primary = header_cells(rows[0])
    header_secondary = header_cells(rows[1]) if len(rows) > 1 else []
    hdr = _merge_header_rows(header_primary, header_secondary)
    qn_idx = find_qn_index_from_header(hdr)
    if qn_idx is None:
        log("[MAIL/DEL] QN 열 인덱스 미발견 → 삭제 불가")
        return table_html, []

    data_start = 2 if header_secondary else 1
    kept = rows[:data_start]
    removed = []
    for r in rows[data_start:]:
        cells = data_cells(r)
        qn = cells[qn_idx] if qn_idx < len(cells) else ""
        if qn.strip() in deleted_qns:
            removed.append(qn.strip())
            continue

        # ✅ 열 개수 보존 (pairs 길이에 맞게 패딩)
        pairs = re.findall(r"(?is)(<t[dh][^>]*>).*?(</t[dh]>)", r)
        filled = cells + [""] * (len(pairs) - len(cells))
        new_r = "<tr>" + "".join(
            f"{pre}{filled[i]}{post}" for i, (pre, post) in enumerate(pairs)
        ) + "</tr>"
        kept.append(new_r)

    new_html = rebuild_table(kept)
    after_rows = count_data_rows(new_html)
    log(f"[MAIL/DEL] 삭제완료 removed={removed} / after_rows={after_rows}")

    residual = [q for q in removed if re.search(rf">{re.escape(q)}<", new_html)]
    if residual:
        log(f"⚠ [MAIL/DEL] 삭제 후에도 남아있는 QN 감지: {residual}")

    return new_html, removed


def append_new_rows_first(table_html, new_rows, header_mapping):
    """첫 번째 표(C1~CK) 신규 행 추가"""
    if not new_rows: 
        log("[MAIL/ADD-FIRST] 추가할 신규 행 없음")
        return table_html

    rows = re.findall(r"(?is)<tr\b.*?>.*?</tr>", table_html)
    if len(rows) < 1:
        log("[MAIL/ADD-FIRST] 표 구조 이상(행 없음)")
        return table_html

    # 헤더 추출
    hdrs = [re.sub(r"(?is)<.*?>","",h).replace("\xa0"," ").strip() 
            for h in re.findall(r"(?is)<t[dh]\b.*?>.*?</t[dh]>", rows[0])]
    hdr_cnt = len(hdrs)
    log(f"[MAIL/ADD-FIRST] 헤더셀수={hdr_cnt}, header_mapping={header_mapping}")

    # --- 매핑 ---
    idx_map = {}
    for key, hdr_sub in header_mapping.items():
        candidates = hdr_sub if isinstance(hdr_sub,(list,tuple)) else [hdr_sub]
        for cand in candidates:
            for i,h in enumerate(hdrs):
                if cand.lower().replace(" ","") in h.lower().replace(" ",""):
                    idx_map[key] = i; break
            if key in idx_map: break
    
    def build_row(values):
        cells = "".join([f"<td style='border:1px solid black;'>{v}</td>" for v in values])
        return f"<tr>{cells}</tr>"


    # --- 신규행 생성 ---
    new_rows_html=[]
    for item in new_rows:
        vals = [""]*hdr_cnt
        for k,v in item.items():
            if k in idx_map:
                hi = idx_map[k]
                if hi < hdr_cnt:
                    vals[hi] = v if v is not None else ""
        # QN_Type 강제 4번째 열 (0-index → 3)
        if "QN_Type" in item and 3 < hdr_cnt:
            vals[3] = item["QN_Type"]
        log(f"[DEBUG] 신규행 데이터(FIRST): {item}")
        new_rows_html.append(build_row(vals))

    final = rows + new_rows_html
    new_html = "<table>" + "".join(final) + "</table>"
    log(f"[MAIL/ADD-FIRST] 신규행 추가 완료: {len(new_rows)}건")
    return new_html

def append_new_rows_second(table_html, new_rows, header_mapping):
    """
    두 번째 표(Q1~Q9) 신규 행 추가 (첫 번째 표 방식 동일)
    - 기존 구조(pairs) 유지
    - 값만 대입
    - 공백행 추가 안 함
    """
    if not new_rows:
        log("[MAIL/ADD-SECOND] 신규 행 없음")
        return table_html

    rows = re.findall(r"(?is)<tr\b.*?>.*?</tr>", table_html)
    if not rows:
        log("[MAIL/ADD-SECOND] 표 구조 없음")
        return table_html

    # 헤더 추출
    hdr_cells = re.findall(r"(?is)<t[dh][^>]*>.*?</t[dh]>", rows[0])
    hdr_cnt = len(hdr_cells)
    log(f"[MAIL/ADD-SECOND] 헤더셀수={hdr_cnt}")

    # 신규행 HTML 저장소
    new_rows_html = []

    for item in new_rows:
        # 완전히 빈 행은 추가하지 않음
        if not any(str(v).strip() for v in item.values()):
            continue

        # 값 매핑
        vals = [""] * hdr_cnt
        for key, val in item.items():
            for i, h in enumerate(hdr_cells):
                norm = lambda x: re.sub(r"[^a-z0-9]", "", x.lower())
                if norm(key) in norm(h):
                    vals[i] = sanitize_cell_value(val)
                    break

        # 기존 셀 구조 복제
        sample_row = rows[-1] if len(rows) > 1 else rows[0]
        pairs = re.findall(r"(?is)(<t[dh][^>]*>).*?(</t[dh]>)", sample_row)
        filled = vals + [""] * (len(pairs) - len(vals))

        new_r = "<tr>" + "".join(
            f"{pre}{filled[i]}{post}" for i, (pre, post) in enumerate(pairs)
        ) + "</tr>"
        new_rows_html.append(new_r)

    # 원래 표 + 신규행 결합
    final = rows + new_rows_html
    new_html = "<table>" + "".join(final) + "</table>"
    log(f"[MAIL/ADD-SECOND] 신규행 추가 완료 ({len(new_rows_html)}건)")
    return new_html


    
# =========================
# QM03 상세조회 (Issue: 첫 표=Attribute, 두 표=Text)
# =========================
TABSTRIP_TOP          = r"wnd[0]/usr/tabsTAB_GROUP_10"
TAB_REFERENCE_OBJECT  = r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01"
TAB_ITEMS             = r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10"
QN_NO_ID              = r"wnd[0]/usr/sub/1/ctxt"
MATERIAL_TXT_DESC     = r"wnd[0]/usr/sub/1/txt[2]"
MATERIAL_NO_DEEP      = r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/ctxt[0]"
TABSTRIP_ITEMS_TAB1   = r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssub/2/3/tabsTAB_GROUP_20/tabp20\TAB01"
ITEMS_SCROLL          = r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssub/2/3/tabsTAB_GROUP_20/tabp20\TAB01/ssub/2/3/1/2"
ITEMS_TABLE_ID        = ITEMS_SCROLL + r"/tblSAPLIQS0POSITION_VIEWER"
ITEMS_ATTR_COL        = 6
ITEMS_TEXT_COL        = 7

def _sel_tab(session, tab_id):
    try: session.findById(tab_id).select(); time.sleep(0.15); return True
    except: return False

def _first_text(session, id_):
    try:
        ctl = session.findById(id_)
        return (getattr(ctl,"text",None) or getattr(ctl,"Text",None) or "").strip() or None
    except: return None

def _get_issue(session, col_idx):
    try: tbl = session.findById(ITEMS_TABLE_ID)
    except: return None
    def _get_cell(r,c):
        try: return tbl.GetCellValue(r,c)
        except: pass
        for kind in ("txt","ctxt","cmb"):
            try:
                cell = session.findById(f"{ITEMS_TABLE_ID}/{kind}[{c},{r}]")
                return getattr(cell,"Text",None) or getattr(cell,"text",None)
            except: pass
        return ""
    for r in (0,1):
        v = _get_cell(r, col_idx)
        vv = ("" if v is None else str(v).strip())
        if vv: return vv
    return None

def _get_issue_attr(session): return _get_issue(session, ITEMS_ATTR_COL)
def _get_issue_text(session): return _get_issue(session, ITEMS_TEXT_COL)

def open_qm03_same_session(session, qn):
    ok = wait_exists(session, "wnd[0]/tbar[0]/okcd", timeout=4)
    if not ok: return False
    try:
        ok.text = "/nQM03"; send_enter(session); time.sleep(0.3)
        fld = wait_exists(session, "/app/con[0]/ses[0]/wnd[0]/usr/ctxt", timeout=3)
        if fld:
            fld.text = str(qn); send_enter(session); time.sleep(0.25)
            if wait_exists(session, TABSTRIP_TOP, timeout=3): return True
    except: return False
    return False

def get_details_for_qns_first(qn_list):
    if not qn_list: return {}
    ses = get_session(); out=[]
    for qn in qn_list:
        log(f"[QM03-First] {qn} 조회")
        if not open_qm03_same_session(ses, qn):
            out.append({"QN":qn,"Material":"","Issue":"","QN_Type":""}); continue
        data={"QN":qn}
        if _sel_tab(ses, TAB_REFERENCE_OBJECT):
            mat_no = _first_text(ses, MATERIAL_NO_DEEP)
            mat_all= _first_text(ses, MATERIAL_TXT_DESC)
            data["Material"] = (mat_no or mat_all) or ""

        # QN Type (ComboBox, Text)
        try:
            qn_type_ctrl = ses.findById("wnd[0]/usr/sub/1/sub/1/1/cmb")
            raw_type = getattr(qn_type_ctrl, "Text", "").strip()
            log(f"[DEBUG] QN TYPE cmb: {raw_type}")
        except:
            try:
                qn_type_ctrl = ses.findById("wnd[0]/usr/sub/1/sub/1/1/txt")
                raw_type = getattr(qn_type_ctrl, "Text", "").strip()
                log(f"[DEBUG] QN TYPE txt: {raw_type}")
            except:
                raw_type = ""
                log("[DEBUG] QN TYPE 찾기 실패")
        
        data["QN_Type"] = simplify_qn_type(raw_type)

        
        # Issue(Attribute) --> items가 없으면 --> description 반환
        if _sel_tab(ses, TAB_ITEMS):
            try: ses.findById(TABSTRIP_ITEMS_TAB1).select()
            except: pass
            data["Issue"] = _get_issue_attr(ses) or ""
        out.append(data)
    return {d["QN"]: d for d in out}

def get_details_for_qns_second(qn_list):
    if not qn_list: return {}
    ses = get_session(); out=[]
    for qn in qn_list:
        log(f"[QM03-Second] {qn} 조회")
        if not open_qm03_same_session(ses, qn):
            log(f"[QM03-Second] 진입 실패 → Material/Issue/QN_Type 빈값")
            out.append({"QN":qn,"Material":"","Issue":"","QN_Type":""}); continue

        data={"QN":qn}

        # Material
        if _sel_tab(ses, TAB_REFERENCE_OBJECT):
            mat_no = _first_text(ses, MATERIAL_NO_DEEP)
            mat_all= _first_text(ses, MATERIAL_TXT_DESC)
            data["Material"] = (mat_no or mat_all) or ""

        # QN Type (ComboBox)
        try:
            qn_type_ctrl = ses.findById("wnd[0]/usr/sub/1/sub/1/1/cmb")
            if qn_type_ctrl:
                raw_type = (getattr(qn_type_ctrl, "Text", "") or "").strip()
                data["QN_Type"] = simplify_qn_type(raw_type)
            else:
                data["QN_Type"] = ""
        except Exception as e:
            log(f"[QM03] QN_Type 추출 실패: {e}")
            data["QN_Type"] = ""

        # Issue (Text 열 → 없으면 Attribute 열 fallback --> 그래도 없으면 description 내용 가져오기)
        issue_val = ""
        if _sel_tab(ses, TAB_ITEMS):
            try:
                ses.findById(TABSTRIP_ITEMS_TAB1).select()
            except:
                pass
            issue_val = _get_issue_text(ses) or _get_issue_attr(ses) or ""

        # items 탭 자체가 없거나 issue 비었을 때 → Description 사용
        if not issue_val:
            try:
                desc_ctrl = ses.findById("wnd[0]/usr/sub/1/txt[2]")
                issue_val = getattr(desc_ctrl, "Text", "") or getattr(desc_ctrl, "text", "")
                issue_val = issue_val.strip()
            except:
                issue_val = ""

        if "2201" in issue_val:
            issue_val = issue_val.replace("2201", "").strip()

        data["Issue"] = issue_val
        out.append(data)
        log(f"[QM03-Second] 결과: {data}")

    return {normalize_qn(d["QN"]): d for d in out}

# =========================
# ZNOT 파싱 (Note number / Task code / Elapsed)
# =========================
def extract_d_values_from_znot(filepath, target_qns):
    if not os.path.isfile(filepath): 
        log(f"[ZNOT] 파일 없음: {filepath}")
        return []
    with open(filepath,"r",encoding="utf-8-sig",errors="ignore") as f:
        lines=[ln.rstrip("\r\n") for ln in f]
    header_idx=None; headers=[]
    for i,ln in enumerate(lines):
        if "Note number" in ln and "Elapsed" in ln:
            header_idx=i
            headers=[h.replace("\xa0"," ").strip() for h in ln.split("\t") if h.strip()]
            break
    if header_idx is None: 
        log("[ZNOT] 헤더 미발견")
        return []
    colmap = {h.lower(): idx for idx,h in enumerate(headers)}
    def get_val(cells, key):
        k = key.lower()
        for name, idx in colmap.items():
            if k in name and idx < len(cells):
                return cells[idx]
        return ""
    target_set = set(str(x) for x in target_qns)
    results=[]
    for ln in lines[header_idx+1:]:
        if not ln.strip(): continue
        raw = ln.split("\t")
        cells=[c.replace("\xa0"," ").strip() for c in raw if c.strip()]
        if not cells: continue
        qn = get_val(cells,"note number")
        if not qn or qn not in target_set: continue
        task = get_val(cells,"task code")
        elapsed = get_val(cells,"elapsed")
        tn = (task or "").upper().replace(" ","")
        if tn in ("D-3C","D-5C","D-8"):
            results.append({"QN":qn,"Task":task,"Elapsed":elapsed})
    log(f"[ZNOT] 추출={len(results)}건 ({results[:]}) / 대상={len(target_set)}건")
    return results
    
def fill_znot_into_first_table(table_html, zmap):
    """기존 첫 번째 표에서 비어 있는 D-3C/5C/8 셀에 zmap 값 채워넣기"""
    rows = table_rows(table_html)
    if len(rows) < 2:
        return table_html

    hdrs = header_cells(rows[0])
    qn_idx = find_qn_index_from_header(hdrs)
    if qn_idx is None:
        return table_html

    def find_col(tag):
        for i, h in enumerate(hdrs):
            if tag.replace(" ", "") in h.upper().replace(" ", ""):
                return i
        return None

    d3_idx = find_col("D-3C")
    d5_idx = find_col("D-5C")
    d8_idx = find_col("D-8")

    new_rows = [rows[0]]
    for r in rows[1:]:
        cells = data_cells(r)
        if not cells:
            new_rows.append(r)
            continue

        qn = cells[qn_idx].strip() if qn_idx < len(cells) else ""
        if not qn:
            new_rows.append(r)
            continue

        z = zmap.get(qn, [])
        def get_elapsed(task_tag):
            for x in z:
                tn = (x.get("Task") or "").upper().replace(" ", "")
                if tn == task_tag:
                    return x.get("Elapsed") or ""
            return ""

        vals = cells[:]
        if d3_idx is not None and d3_idx < len(vals) and not vals[d3_idx]:
            vals[d3_idx] = get_elapsed("D-3C")
        if d5_idx is not None and d5_idx < len(vals) and not vals[d5_idx]:
            vals[d5_idx] = get_elapsed("D-5C")
        if d8_idx is not None and d8_idx < len(vals) and not vals[d8_idx]:
            vals[d8_idx] = get_elapsed("D-8")

        pairs = re.findall(r"(?is)(<t[dh][^>]*>).*?(</t[dh]>)", r)
        cnt = min(len(vals), len(pairs))
        new_r = "<tr>" + "".join(
            f"{pairs[i][0]}{vals[i]}{pairs[i][1]}" for i in range(cnt)
        ) + "</tr>"

        new_rows.append(new_r)

    return rebuild_table(new_rows)
    
# =========================
# No. 열 재정렬 유틸 (첫 표/두 표 별도)
# =========================
def _renumber_table_core(table_html, data_start):
    rows = table_rows(table_html)
    if len(rows) <= data_start + 1:
        return table_html

    no_idx = next((i for i, h in enumerate(header_cells(rows[0])) if "no" in h.lower()), None)
    if no_idx is None:
        return table_html

    new_rows = rows[:data_start]
    for num, r in enumerate(rows[data_start:], start=1):
        cells = data_cells(r)
        if not cells:
            continue
        pairs = re.findall(r"(?is)(<t[dh][^>]*>).*?(</t[dh]>)", r)
        vals = cells[:]
        while len(vals) < no_idx + 1:
            vals.append("")
        vals[no_idx] = str(num)
        new_r = "<tr>" + "".join(
            f"{pre}{vals[i] if i < len(vals) else ''}{post}"
            for i, (pre, post) in enumerate(pairs)
        ) + "</tr>"
        new_rows.append(new_r)

    return rebuild_table(new_rows)

def renumber_first_table(table_html):
    """첫 번째 표 (C1~CK): 데이터 시작 행 = 3번째 (data_start=2)"""
    return _renumber_table_core(table_html, data_start=2)

def renumber_second_table(table_html):
    """두 번째 표 (Q1~Q9): 데이터 시작 행 = 4번째 (data_start=3)"""
    return _renumber_table_core(table_html, data_start=3)

# =========================
# outlook 메일 표 텍스트 가운데 정렬, 폰트 사이즈:10, 폰트 Aptos
# =========================
# Word 정렬 상수
wdAlignParagraphCenter = 1
wdAlignVerticalCenter  = 1

def apply_final_format_to_tables():
    """표 내용 전체를 한번에 가운데 정렬 + 맑은 고딕 10pt로 통일 (고속 적용)"""
    ol = win32.Dispatch("Outlook.Application")
    inspector = ol.ActiveInspector()
    if not inspector:
        raise RuntimeError("열려 있는 Outlook 메일창이 없습니다.")

    doc = inspector.WordEditor
    tables = doc.Tables
    if tables.Count < 1:
        log("⚠ 표가 없습니다. 본문 전체 서식은 건너뜁니다.")
        return

    for i in range(1, tables.Count + 1):
        tbl = tables.Item(i)
        rng = tbl.Range

        # 🔹 표 전체 폰트 한번에 적용
        rng.Font.Name = "맑은 고딕"
        rng.Font.Size = 10

        # 🔹 가로 가운데 정렬 한번에 적용
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter

        # 🔹 세로 가운데 정렬은 표 전체 속성으로 설정
        tbl.Rows.Alignment = 0  # 표 자체는 왼쪽 정렬 유지
        tbl.Range.Cells.VerticalAlignment = wdAlignVerticalCenter

    log("✅ 표 전체 단위로 가운데 정렬 + 맑은 고딕 10pt 적용 완료 (고속 모드)")


# =========================
# Outlook 표 스타일 적용 ("표 눈금 밝게")
# =========================
def apply_table_nungum_style_to_all_tables():
    """
    Word(Outlook) 표 스타일 중 '표 눈금 밝게' (Table Grid Light)를
    모든 표에 일괄 적용.
    """
    ol = win32.Dispatch("Outlook.Application")
    inspector = ol.ActiveInspector()
    if not inspector:
        raise RuntimeError("열려 있는 Outlook 메일창이 없습니다.")
    doc = inspector.WordEditor

    # 스타일 이름 후보 (한글/영문 혼용 환경 대응)
    candidates = ["표 눈금 밝게", "Table Grid Light", "Table Grid", "표 격자"]

    # 스타일 이름 탐색
    def _resolve_style(doc, names):
        for n in names:
            try:
                _ = doc.Styles(n)
                return n
            except:
                try:
                    _ = doc.Styles(n).NameLocal
                    return n
                except:
                    continue
        return None

    style_name = _resolve_style(doc, candidates)
    if not style_name:
        log("⚠ 표 스타일 '표 눈금 밝게'를 찾지 못해 기본 유지.")
        return

    app = doc.Application
    try:
        prev = app.ScreenUpdating
        app.ScreenUpdating = False
    except:
        prev = True

    count = 0
    for t in doc.Tables:
        try:
            t.Style = style_name
            count += 1
        except Exception as e:
            log(f"[WARN] 표 스타일 적용 실패: {e}")

    try:
        app.ScreenUpdating = prev
    except:
        pass

    log(f"✅ '{style_name}' 스타일 {count}개 표에 적용 완료.")

#색상 조건            
def rgb(r,g,b):
    return r + (g << 8) + (b << 16)

def apply_table_color_rules(data_start_row=3):
    """첫 번째 표에서 Elapsed time 행 제외 후
       D-3C/D-5C/D-8 은 조건부 색상 적용,
       '이슈 내용'(6번째 열)은 무조건 색상 제거
    """
    ol = win32.Dispatch("Outlook.Application")
    inspector = ol.ActiveInspector()
    if not inspector:
        raise RuntimeError("열려있는 Outlook 메일창이 없습니다.")

    doc = inspector.WordEditor
    tables = doc.Tables
    if tables.Count < 1:
        raise RuntimeError("메일에 표가 없습니다.")

    tbl = tables.Item(1)

    # 헤더 텍스트 읽기
    headers = [tbl.Cell(1, c).Range.Text.strip().replace("\r"," ").replace("\x07"," ") 
               for c in range(1, tbl.Columns.Count+1)]

    idx_d3  = next((i+1 for i,h in enumerate(headers) if "D-3C" in h.upper()), None)
    idx_d5  = next((i+1 for i,h in enumerate(headers) if "D-5C" in h.upper()), None)
    idx_d8  = next((i+1 for i,h in enumerate(headers) if "D-8"  in h.upper()), None)

    green_bg = rgb(226,239,217)
    red_bg   = rgb(251,228,213)

    issue_col = 6  # '이슈 내용' 열 고정

    start_row = max(2, data_start_row)

    # 데이터 시작 행 (1=헤더, 2=Elapsed time)
    for r in range(start_row, tbl.Rows.Count+1):
        # --- 조건부 색상: D-3C, D-5C, D-8 ---
        for (idx, threshold) in [(idx_d3,1), (idx_d5,5), (idx_d8,20)]:
            if idx:
                cell = tbl.Cell(r, idx)
                val = cell.Range.Text.strip().replace("\r"," ").replace("\x07"," ")
                cell.Shading.BackgroundPatternColor = green_bg
                try:
                    if float(val) >= threshold:
                        cell.Shading.BackgroundPatternColor = red_bg
                except:
                    pass

        # --- 이슈 내용 열은 무조건 흰색으로 덮어쓰기 ---
        if issue_col <= tbl.Columns.Count:
            try:
                tbl.Cell(r, issue_col).Shading.BackgroundPatternColor = 16777215
            except Exception:
                pass

    log("✅ 색상 규칙 적용 완료 (D-3C/5C/8 조건부, 이슈내용 열 색상 제거)")


def clear_table_header_emphasis():
    """모든 표의 헤더 행에 적용된 강조 서식을 제거."""
    ol = win32.Dispatch("Outlook.Application")
    inspector = ol.ActiveInspector()
    if not inspector:
        raise RuntimeError("열려있는 Outlook 메일창이 없습니다.")

    doc = inspector.WordEditor
    for tbl in doc.Tables:
        try:
            header_row = tbl.Rows(1)
        except Exception:
            continue

        try:
            rng = header_row.Range
            rng.Font.Bold = False
            rng.Font.ColorIndex = 0  # 자동 색상
            header_row.Shading.BackgroundPatternColor = 16777215  # 흰색
        except Exception as exc:
            log(f"[WARN] 헤더 강조 제거 실패: {exc}")


def apply_final_table_styles_then_colors(
    data_start_row=3,
    apply_header_style=True,
    clear_header_emphasis=True,
):
    """
    1️⃣ (선택) 모든 표에 '표 눈금 밝게' 스타일 적용
    2️⃣ 첫 번째 표에 색상 규칙 적용
    3️⃣ 모든 표에 가운데정렬·폰트·사이즈 적용
    4️⃣ 본문 전체 폰트도 맑은 고딕 10pt로 통일
    """
    if apply_header_style:
        apply_table_nungum_style_to_all_tables()
        if clear_header_emphasis:
            clear_table_header_emphasis()

    apply_table_color_rules(data_start_row)
    apply_final_format_to_tables()
    try:
        apply_uniform_body_font()
    except Exception:
        log_exc("UNIFORM_FONT_GLOBAL")


def apply_uniform_body_font(mail_item=None, font_name="맑은 고딕", font_size=10):
    """메일 본문 전체에 통일된 폰트/크기를 적용."""
    try:
        _, doc = _ensure_word_editor(mail_item)
        rng = doc.Content
        rng.Font.Name = font_name
        rng.Font.Size = font_size
        log(f"✅ 본문 전체 폰트를 '{font_name}' {font_size}pt로 통일했습니다.")
    except RuntimeError as exc:
        if mail_item is None:
            log(f"[WARN] 본문 폰트 통일 실패 (Inspector 없음): {exc}")
            return
        try:
            html = mail_item.HTMLBody or ""
            wrapper = f"<div style=\"font-family:'{font_name}'; font-size:{font_size}pt;\">"
            mail_item.HTMLBody = f"{wrapper}{html}</div>"
            log(f"[WARN] WordEditor 없음 → HTML 덮어쓰기 방식으로 폰트 적용 완료")
        except Exception as inner_exc:
            log(f"[ERROR] 본문 폰트 HTML 적용 실패: {inner_exc}")


# -----------------------------------------
# 1️⃣ 셀 내부 값 정리 (숨은 공백/개행/NBSP 제거)
# -----------------------------------------
_WS_ZEROWIDTH = ("\u200b", "\u200c", "\u200d", "\ufeff")

def sanitize_cell_value(v: object) -> str:
    """셀 내부 문자열에서 Outlook이 오인식하는 공백/개행/제로폭 제거"""
    s = "" if v is None else str(v)
    for z in _WS_ZEROWIDTH:
        s = s.replace(z, "")
    s = (
        s.replace("\xa0", " ")   # NBSP → 일반 space
         .replace("\t", " ")
         .replace("\r", "")
         .replace("\n", "")
    )
    return s.strip()


# -----------------------------------------
# 2️⃣ Outlook HTML 최종 정리
# -----------------------------------------
def finalize_outlook_html(new_html: str) -> str:
    """
    Outlook용 HTML 최종 정리:
    - table-layout 고정 (자동 폭 계산 방지)
    - 태그 사이 불필요한 공백 제거
    """
    def lock_table(tag: str) -> str:
        if 'style=' in tag:
            return re.sub(
                r'style=["\']',
                "style='table-layout:fixed;border-collapse:collapse; ",
                tag,
                count=1
            )
        return tag.replace("<table", "<table style='table-layout:fixed;border-collapse:collapse'")
    new_html = re.sub(r"(?is)<table\b[^>]*>", lambda m: lock_table(m.group(0)), new_html)
    new_html = re.sub(r">\s+<", "><", new_html)
    return new_html


# -----------------------------------------
# 3️⃣ 첫 번째/두 번째 표 값만 정리 (스타일 보존)
# -----------------------------------------
def sanitize_table_cells_preserve_structure(table_html: str) -> str:
    """<td>/<th> 내부 텍스트만 정리, HTML 구조와 스타일은 그대로 유지"""
    pattern = r"(?is)(<t[dh][^>]*>)(.*?)(</t[dh]>)"

    def clean_cell(m):
        open_tag, content, close_tag = m.groups()
        clean_text = sanitize_cell_value(re.sub(r"(?is)<.*?>", "", content))
        return f"{open_tag}{clean_text}{close_tag}"

    return re.sub(pattern, clean_cell, table_html)


# -----------------------------------------
# 4️⃣ 두 번째 표(Q1~Q9) 전용 안정화 보정
# -----------------------------------------
def sanitize_table_second(table_html: str) -> str:
    """
    두 번째 표(Q1~Q9) 전용 정리 (최종 안정화 버전)
    - 셀 내부 텍스트의 숨은 공백/제로폭 문자 제거
    - 실제 헤더 열 수에 맞춰 정확히 <td> 개수 유지
    - 초과 셀은 버리고, 부족하면 <td></td>로 채움 (보더 없음)
    - 기존 스타일/링크/정렬은 유지
    """
    if not table_html:
        return table_html

    # 1️⃣ 셀 내부 텍스트 정리
    pattern = r"(?is)(<t[dh][^>]*>)(.*?)(</t[dh]>)"
    def clean_cell(m):
        open_tag, content, close_tag = m.groups()
        clean_text = sanitize_cell_value(re.sub(r"(?is)<.*?>", "", content))
        return f"{open_tag}{clean_text}{close_tag}"
    cleaned_html = re.sub(pattern, clean_cell, table_html)

    # 2️⃣ 행 추출
    rows = re.findall(r"(?is)<tr\b.*?>.*?</tr>", cleaned_html)
    if not rows:
        return cleaned_html
    
    # 3️⃣ 헤더 기준 
    header_cells = [
        c for c in re.findall(r"(?is)<t[dh]\b[^>]*>.*?</t[dh]>", rows[0])
        if strip_html(c).strip() not in ("", "&nbsp;")
    ]
    hdr_cnt = len(header_cells)
    if hdr_cnt == 0:
        # ✅ 혹시 헤더가 비어있다면, 두 번째 행에서라도 기준 잡기
        for r in rows[1:]:
            tmp = re.findall(r"(?is)<t[dh]\b[^>]*>.*?</t[dh]>", r)
            if tmp:
                header_cells = tmp
                hdr_cnt = len(tmp)
                break
    if hdr_cnt == 0:
        return cleaned_html

    fixed_rows = []
    for i, r in enumerate(rows):
        cells = re.findall(r"(?is)<t[dh]\b[^>]*>.*?</t[dh]>", r)
        c_count = len(cells)

        # 헤더 행은 그대로 유지
        if i == 0:
            fixed_rows.append(r)
            continue

        # 데이터 행만 정규화
        if c_count < hdr_cnt:
            # 부족하면 <td></td>로 채움
            for _ in range(hdr_cnt - c_count):
                cells.append("<td></td>")
        elif c_count > hdr_cnt:
            # ✅ 초과 셀은 버리되 로그 남기기
            log(f"[WARN] 행 {i}에 헤더보다 {c_count - hdr_cnt}개 많은 셀 감지 → 잘라냄")
            cells = cells[:hdr_cnt]

        fixed_rows.append("<tr>" + "".join(cells) + "</tr>")

    fixed_html = "<table>" + "".join(fixed_rows) + "</table>"
    fixed_html = finalize_outlook_html(fixed_html)
    return fixed_html

def normalize_qn(qn):
    """QN 값 내 숨은 공백, 제로폭 문자 제거 + 문자열 통일"""
    if qn is None:
        return ""
    s = str(qn)
    for z in ("\u200b", "\u200c", "\u200d", "\ufeff", "\xa0"):
        s = s.replace(z, "")
    return s.strip()

def safe_get(d, *keys):
    """키 이름 불일치 방지를 위한 안전 접근기"""
    if not d:
        return ""
    # 우선 순위대로 직접 시도
    for k in keys:
        if k in d:
            return d[k]
    # 정규화된 키 매칭 (공백/언더바/대소문자 무시)
    dn = {re.sub(r"[^a-z0-9]", "", k.lower()): v for k, v in d.items()}
    for k in keys:
        nk = re.sub(r"[^a-z0-9]", "", k.lower())
        if nk in dn:
            return dn[nk]
    return ""

def _norm_header(s: str) -> str:
    import re
    s = "" if s is None else str(s)
    s = re.sub(r"[\u200b\u200c\u200d\ufeff\xa0]", "", s)  # 제로폭/NBSP 제거
    return re.sub(r"\s+", "", s).lower()  # 공백 제거 + 소문자


def build_first_mail_from_files(day):
    """보낸편지함에 QN 메일이 없을 때, SAP Export 파일만으로 초기 메일 생성 (본문=표 2개만)"""
    log(f"[INIT] 초기 메일 빌드 시작 ({day})")

    # 1) 필요한 파일 확보
    ensure_files_for_test_date(day)
    f_first  = path_first(day)
    f_second = path_second(day)
    f_znot   = path_znot(day)

    # 2) 파싱
    rows_first  = parse_export_txt(f_first,  anchor_qn=None) if os.path.isfile(f_first)  else []
    rows_second = parse_export_txt(f_second, anchor_qn=None) if os.path.isfile(f_second) else []

    # 3) 상세조회 (기존 함수 사용)
    detail_map_first  = get_details_for_qns_first([r["QN"] for r in rows_first])   if rows_first  else {}
    detail_map_second = get_details_for_qns_second([r["QN"] for r in rows_second]) if rows_second else {}

    # 4) 헤더(누락 보완) + 공통 셀 스타일
    hdrs_first  = ["No.", "등록일", "QN #", "QN TYPE", "품번", "이슈 내용", "D-3C", "D-5C", "D-8"]
    hdrs_second = ["No.", "등록일", "QN #", "QN TYPE", "품번", "이슈 내용"]
    base_style = "font-family:'Malgun Gothic'; font-size:10pt; text-align:center;"

    header_html_first  = "".join(f"<th style='{base_style}'>{h}</th>" for h in hdrs_first)
    header_html_second = "".join(f"<th style='{base_style}'>{h}</th>" for h in hdrs_second)

    # 5) 첫 번째 표 행
    rows_html_first = []
    for i, r in enumerate(rows_first, start=1):
        qn = r.get("QN", "")
        d  = detail_map_first.get(qn, {})
        vals = [
            str(i),
            sanitize_cell_value(r.get("Created", "")),
            sanitize_cell_value(qn),
            sanitize_cell_value(d.get("QN_Type", "")),
            sanitize_cell_value(d.get("Material", "")),
            sanitize_cell_value(d.get("Issue", "")),
            "", "", "",  # D-3C/5C/8은 ZNOT로 채움
        ]
        rows_html_first.append("<tr>" + "".join(f"<td style='{base_style}'>{v}</td>" for v in vals) + "</tr>")

    table_first = f"""
    <table style='border-collapse:collapse; table-layout:fixed; border:1px solid #ccc;
                  margin-left:auto; margin-right:auto; font-family:"Malgun Gothic"; font-size:10pt;'>
        <tr>{header_html_first}</tr>
        {''.join(rows_html_first)}
    </table>
    """

    # 5.5) ZNOT 보강 (기존 로직 재사용)
    try:
        if os.path.isfile(f_znot) and rows_first:
            qns_first = [r["QN"] for r in rows_first]
            z_rows = extract_d_values_from_znot(f_znot, qns_first)
            zmap = {}
            for zr in z_rows:
                zmap.setdefault(zr["QN"], []).append({
                    "Task": zr.get("Task", ""),
                    "Elapsed": zr.get("Elapsed", "")
                })
            table_first = fill_znot_into_first_table(table_first, zmap)
            log(f"[INIT] ZNOT 병합 완료: 대상={len(qns_first)}건 / 추출={len(z_rows)}건")
    except Exception as e:
        log(f"[INIT][WARN] ZNOT 병합 중 오류: {e}")

    # 6) 두 번째 표 행
    rows_html_second = []
    for i, r in enumerate(rows_second, start=1):
        qn = r.get("QN", "")
        d  = detail_map_second.get(qn, {})
        vals = [
            str(i),
            sanitize_cell_value(r.get("Created", "")),
            sanitize_cell_value(qn),
            sanitize_cell_value(d.get("QN_Type", "")),
            sanitize_cell_value(d.get("Material", "")),
            sanitize_cell_value(d.get("Issue", "")),
        ]
        rows_html_second.append("<tr>" + "".join(f"<td style='{base_style}'>{v}</td>" for v in vals) + "</tr>")

    table_second = f"""
    <table style='border-collapse:collapse; table-layout:fixed; border:1px solid #ccc;
                  margin-top:10px; margin-left:auto; margin-right:auto;
                  font-family:"Malgun Gothic"; font-size:10pt;'>
        <tr>{header_html_second}</tr>
        {''.join(rows_html_second)}
    </table>
    """

    # 7) 메일 생성 (초기 표 2개 포함)
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = f"{day.month}/{day.day} QN 등록 현황 공유"
    recipients = MAIL_RECIPIENTS if MAIL_RECIPIENTS else list(DEFAULT_MAIL_RECIPIENTS)
    mail.To = "; ".join(recipients)
    mail.HTMLBody = f"<html><body>{table_first}<br><br>{table_second}</body></html>"
    mail.Display()
    try:
        insp = mail.GetInspector()
        insp.Display()
        insp.Activate()
    except Exception:
        log("[INIT][WARN] Outlook Inspector 활성화 실패")

    # 8) 스타일 적용 (기존 함수 사용)
    try:
        apply_table_nungum_style_to_all_tables()
        apply_table_color_rules_initial()
        apply_final_format_to_tables()
        log("✅ [INIT] 초기 전용 스타일(눈금→색상→정렬) 적용 완료")
    except Exception as e:
        log(f"[INIT] 스타일 적용 중 오류: {e}")

    # 9) 본문 요약 및 KPI 블록 작성
    try:
        mail_obj, total, sat_d3, sat_d5, sat_d8 = parse_first_table_C1Ck_2row_header(mail)
        append_summary_in_mail(mail_obj, total, sat_d3, sat_d5, sat_d8)
        log("✅ [INIT] 본문 요약 및 KPI 블록 추가 완료")
    except Exception as e:
        log(f"[INIT][WARN] 본문 요약 추가 중 오류: {e}")

    try:
        apply_uniform_body_font(mail)
        log("✅ [INIT] 본문 폰트/크기 통일 완료")
    except Exception as e:
        log(f"[INIT][WARN] 본문 폰트 통일 실패: {e}")

    return mail



def apply_table_color_rules_initial():
    """첫 번째 표에서 2행(헤더 바로 아래)부터
       D-3C/D-5C/D-8 은 조건부 색상 적용,
       '이슈 내용'(6번째 열)은 무조건 흰색으로 초기화
    """
    ol = win32.Dispatch("Outlook.Application")
    inspector = ol.ActiveInspector()
    if not inspector:
        raise RuntimeError("열려있는 Outlook 메일창이 없습니다.")

    doc = inspector.WordEditor
    tables = doc.Tables
    if tables.Count < 1:
        raise RuntimeError("메일에 표가 없습니다.")

    tbl = tables.Item(1)

    # --- 헤더 읽기 ---
    headers = [
        tbl.Cell(1, c).Range.Text.replace("\r", "").replace("\x07", "").strip()
        for c in range(1, tbl.Columns.Count + 1)
    ]

    # --- 열 인덱스 찾기 ---
    def find_col(key):
        key = key.upper().replace(" ", "")
        for i, h in enumerate(headers):
            if key in h.upper().replace(" ", ""):
                return i + 1
        return None

    idx_d3 = find_col("D-3C")
    idx_d5 = find_col("D-5C")
    idx_d8 = find_col("D-8")
    issue_col = 6  # '이슈 내용' 열

    # --- 색상 정의 ---
    def rgb(r, g, b): return r + (g << 8) + (b << 16)
    green_bg = rgb(226, 239, 217)
    red_bg = rgb(251, 228, 213)
    white_bg = rgb(255, 255, 255)

    # ✅ 데이터 시작 행 = 2 (헤더 바로 아래부터)
    for r in range(2, tbl.Rows.Count + 1):
        # 조건부 색상: D-3C / D-5C / D-8
        for (idx, threshold) in [(idx_d3, 1), (idx_d5, 5), (idx_d8, 20)]:
            if idx:
                cell = tbl.Cell(r, idx)
                val = cell.Range.Text.replace("\r", "").replace("\x07", "").strip()
                cell.Shading.BackgroundPatternColor = green_bg
                try:
                    if float(val) >= threshold:
                        cell.Shading.BackgroundPatternColor = red_bg
                except:
                    pass

        # '이슈 내용' 열은 무조건 흰색으로 덮기
        if issue_col <= tbl.Columns.Count:
            try:
                tbl.Cell(r, issue_col).Shading.BackgroundPatternColor = white_bg
            except:
                pass

    log("✅ [INIT] 색상 규칙 적용 완료 (2행부터, D-3C/5C/8 조건부 + 이슈내용 열 흰색)")


#=============================================================통계 및 본문내용 작성====================================================
def safe_float(s):
    """문자열에서 숫자만 추출해 float으로 변환"""
    if not s:
        return None
    try:
        s = str(s).strip().replace(",", ".")
        return float(re.sub(r"[^0-9.]", "", s))
    except:
        return None


def _ensure_word_editor(mail_item=None):
    """해당 MailItem의 Word 편집기를 확보한다 (필요 시 Display 호출)."""
    outlook = win32.Dispatch("Outlook.Application")
    inspector = None
    target_mail = mail_item

    def _display_mail(item):
        if not item:
            return
        try:
            item.Display()
        except Exception:
            pass

    if target_mail is not None:
        _display_mail(target_mail)
        try:
            inspector = target_mail.GetInspector()
        except Exception:
            inspector = None

    if inspector is None:
        inspector = outlook.ActiveInspector()
        if inspector is None:
            raise RuntimeError("열려 있는 Outlook 메일창이 없습니다.")

    for _ in range(3):
        try:
            doc = inspector.WordEditor
            return inspector, doc
        except AttributeError:
            if target_mail is not None:
                try:
                    # HTML 편집기로 전환 시도
                    if hasattr(target_mail, "BodyFormat"):
                        target_mail.BodyFormat = 2  # olFormatHTML
                except Exception:
                    pass
                _display_mail(target_mail)
                try:
                    inspector = target_mail.GetInspector()
                except Exception:
                    pass
            else:
                try:
                    inspector.Activate()
                except Exception:
                    pass

    raise RuntimeError("Outlook Word 편집기를 초기화하지 못했습니다.")


###############################################################################
# Outlook 메일의 첫 번째 표(C1~Ck, 2줄 헤더 구조) 통계 계산
###############################################################################
def parse_first_table_C1Ck_2row_header(mail_obj=None):
    """C1~Ck 표 전체 데이터를 HTML 기반으로 스캔해 KPI 통계를 계산한다."""
    mail = mail_obj
    if mail is None:
        outlook = win32.Dispatch("Outlook.Application")
        insp = outlook.ActiveInspector()
        if insp is None:
            raise RuntimeError("현재 열려 있는 Outlook 메일이 없습니다.")
        mail = insp.CurrentItem

    html = mail.HTMLBody or ""
    tables = get_tables(html)
    if not tables:
        raise RuntimeError("메일 본문에서 표를 찾지 못했습니다.")

    def _norm(text):
        return re.sub(r"\s+", "", (text or "").strip()).lower()

    def _combine_headers(row_primary, row_secondary):
        max_len = max(len(row_primary), len(row_secondary))
        headers_local = []
        for i in range(max_len):
            top = row_primary[i] if i < len(row_primary) else ""
            fallback = row_secondary[i] if i < len(row_secondary) else ""
            headers_local.append(top or fallback)
        return headers_local

    def _find_col(headers_local, keywords):
        for i, h in enumerate(headers_local):
            H = _norm(h)
            if any(k in H for k in keywords):
                return i
        return None

    selected_tbl = None
    headers = []
    selected_idx = None
    parsed_rows = []

    for idx, tbl_html in enumerate(tables, start=1):
        rows = table_rows(tbl_html)
        if len(rows) < 2:
            continue
        header_primary = header_cells(rows[0])
        header_secondary = header_cells(rows[1]) if len(rows) > 1 else []
        headers_candidate = _combine_headers(header_primary, header_secondary)

        idx_qn = _find_col(headers_candidate, ["qn", "notification", "qn#"])
        idx_d3 = _find_col(headers_candidate, ["d-3c", "d3c", "d-3"])
        idx_d5 = _find_col(headers_candidate, ["d-5c", "d5c", "d-5"])
        idx_d8 = _find_col(headers_candidate, ["d-8", "d8", "d-8c"])

        if all(v is not None for v in [idx_qn, idx_d3, idx_d5, idx_d8]):
            selected_tbl = tbl_html
            headers = headers_candidate
            selected_idx = idx
            parsed_rows = rows
            break

    if selected_tbl is None:
        raise RuntimeError("필수 열(QN, D-3C, D-5C, D-8)을 가진 표를 찾지 못했습니다.")

    print(f"[INFO] KPI 대상 표: index={selected_idx} | 행수={len(parsed_rows)}")
    print(f"[INFO] 헤더 감지: {headers}")

    idx_qn = _find_col(headers, ["qn", "notification", "qn#"])
    idx_d3 = _find_col(headers, ["d-3c", "d3c", "d-3"])
    idx_d5 = _find_col(headers, ["d-5c", "d5c", "d-5"])
    idx_d8 = _find_col(headers, ["d-8", "d8", "d-8c"])

    total = 0
    sat_d3 = sat_d5 = sat_d8 = 0

    second_row_text = "".join(_norm(text) for text in data_cells(parsed_rows[1])) if len(parsed_rows) > 1 else ""
    data_start_row = 2 if "elapsed" in second_row_text else 1

    for row_idx in range(data_start_row, len(parsed_rows)):
        try:
            cells = data_cells(parsed_rows[row_idx])
            row_payload = "".join(cells).strip()
            if not row_payload:
                continue

            qn = cells[idx_qn] if idx_qn < len(cells) else ""
            val_d3 = safe_float(cells[idx_d3] if idx_d3 < len(cells) else "")
            val_d5 = safe_float(cells[idx_d5] if idx_d5 < len(cells) else "")
            val_d8 = safe_float(cells[idx_d8] if idx_d8 < len(cells) else "")

            if (qn or "").lower() == "elapsed time":
                continue

            total += 1
            qn_label = qn or f"ROW-{row_idx - data_start_row + 1}"
            print(f"[ROW] QN={qn_label} | D-3C={val_d3} | D-5C={val_d5} | D-8={val_d8}")

            if val_d3 is not None and val_d3 < 1:
                sat_d3 += 1
            if val_d5 is not None and val_d5 < 5:
                sat_d5 += 1
            if val_d8 is not None and val_d8 < 20:
                sat_d8 += 1
        except Exception as e:
            print(f"[WARN] {row_idx}행 처리 실패: {e}")
            continue

    print("\n=== 통계 요약 ===")
    print(f"총 건수: {total}")
    print(f"D-3C 만족: {sat_d3} / {total}")
    print(f"D-5C 만족: {sat_d5} / {total}")
    print(f"D-8 만족: {sat_d8} / {total}")

    return mail, total, sat_d3, sat_d5, sat_d8


def append_summary_in_mail(mail, total, sat_d3, sat_d5, sat_d8):
    html = mail.HTMLBody or ""

    safe_total = total if total else 0
    def _pct(num, denom):
        return (num / denom * 100) if denom else 0

    pct_d3 = _pct(sat_d3, safe_total)
    pct_d5 = _pct(sat_d5, safe_total)
    pct_d8 = _pct(sat_d8, safe_total)

    # ────────────────────────────────
    # 모든 표 추출 (보존)
    # ────────────────────────────────
    tables = re.findall(r"(?is)(<table.*?</table>)", html)

    # ────────────────────────────────
    # 상단 신규 블록 생성
    # ────────────────────────────────
    date_str = f"{TEST_DATE.year}년 {TEST_DATE.month}월 {TEST_DATE.day}일"
    header_html = (
        f"<div style='font-family:\"Malgun Gothic\"; font-size:10pt; margin:6px 0;'>"
        f"{date_str} QN 등록 현황 공유 드립니다.<br>"
        f"현재까지 QN은 총 <b>{safe_total}</b>건입니다."
        f"</div><br>"
        f"<div style='font-family:\"Malgun Gothic\"; font-size:10pt; margin:6px 0;'>"
        f"<b>QN 입력 시간 분석 결과</b><br>"
        f"D-3C (Within 1day): {sat_d3}건 / {safe_total}건 ({pct_d3:.1f}%)<br>"
        f"D-5C (Within 5days): {sat_d5}건 / {safe_total}건 ({pct_d5:.1f}%)<br>"
        f"D-8 (Within 20days): {sat_d8}건 / {safe_total}건 ({pct_d8:.1f}%)"
        f"</div><hr>"
    )

    # ────────────────────────────────
    # 표 위/아래 문장 구성
    # ────────────────────────────────
    new_body = header_html

    if len(tables) >= 1:
        new_body += (
            "<p style='font-family:\"Malgun Gothic\"; font-size:12pt; font-weight:bold;'>"
            "Notification Type C1-CK</p>" + tables[0]
        )
    if len(tables) >= 2:
        new_body += (
            "<p style='font-family:\"Malgun Gothic\"; font-size:12pt; font-weight:bold;'>"
            "Notification Type Q1-Q9</p>" + tables[1] +
            "<p style='font-family:\"Malgun Gothic\"; font-size:10pt;'>이상입니다.</p>"
        )

    # 나머지 표 (3번째 이후)는 원래 순서대로 유지
    if len(tables) > 2:
        for t in tables[2:]:
            new_body += t


    # ────────────────────────────────
    # 메일 본문 반영 + 폰트 통일
    # ────────────────────────────────
    wrapper_style = "font-family:'Malgun Gothic'; font-size:10pt;"
    mail.HTMLBody = f"<div style=\"{wrapper_style}\">{new_body}</div>"
    try:
        apply_uniform_body_font(mail)
    except Exception:
        log_exc("UNIFORM_FONT_SUMMARY")


def _normalize_header_for_anchor(text):
    """헤더 텍스트를 앵커 검색용으로 정규화."""
    raw = strip_html(text or "")
    cleaned = re.sub(r"[\s\xa0]+", "", raw)
    return cleaned.upper()


def _merge_header_rows(primary, secondary):
    if not primary and not secondary:
        return []
    if not secondary:
        return primary
    merged = []
    max_len = max(len(primary), len(secondary))
    for i in range(max_len):
        top = primary[i] if i < len(primary) else ""
        bottom = secondary[i] if i < len(secondary) else ""
        merged.append(top or bottom)
    return merged


def _classify_table_for_anchor(table_html):
    """표가 C-Type(첫 번째)인지 Q-Type(두 번째)인지 판별."""
    rows = table_rows(table_html)
    if not rows:
        return None

    primary = header_cells(rows[0])
    secondary = header_cells(rows[1]) if len(rows) > 1 else []
    headers = _merge_header_rows(primary, secondary)
    if not headers:
        return None

    if find_qn_index_from_header(headers) is None:
        return None

    norm_headers = [_normalize_header_for_anchor(h) for h in headers]
    has_d_cols = any(
        any(tag in h for tag in ("D-3", "D3", "D-5", "D5", "D-8", "D8"))
        for h in norm_headers
    )

    if has_d_cols:
        return "first"
    return "second"


def detect_anchor_tables(tables):
    """메일 본문에서 C/Q 표를 찾아 인덱스와 HTML을 반환."""
    if len(tables) < 2:
        raise RuntimeError("메일에 표가 2개 미만입니다.")

    first_idx = second_idx = None
    for idx, tbl in enumerate(tables):
        kind = _classify_table_for_anchor(tbl)
        if kind == "first" and first_idx is None:
            first_idx = idx
        elif kind == "second" and second_idx is None:
            second_idx = idx

        if first_idx is not None and second_idx is not None:
            break

    fallback_used = False
    if first_idx is None:
        first_idx = 0
        fallback_used = True
    if second_idx is None:
        for idx in range(len(tables)):
            if idx != first_idx:
                second_idx = idx
                fallback_used = True
                break

    if second_idx is None:
        raise RuntimeError("메일에서 Q-Type 표를 찾지 못했습니다.")

    if fallback_used:
        log("[WARN] 표 식별 실패로 기본 순서를 사용했습니다.")

    return first_idx, tables[first_idx], second_idx, tables[second_idx]

# =========================
# 메인 파이프라인 (TEST_DATE 기준)
# =========================
# GUI에서 진행 상태를 표시할 때 사용되는 단계별 안내 메시지
STEP_MESSAGES = {
    0: "SAP 트랜잭션에서 오늘 날짜 기준 전체 QN 데이터를 조회하고 있습니다.",
    1: "기존에 발송했던 보고서 메일을 검색 중입니다...",
    2: "새로 등록된 C-Type QN 항목을 분석 중입니다...",
    3: "삭제된 C-Type QN 항목을 분석 중입니다...",
    4: "새로 등록된 Q-Type QN 항목을 분석 중입니다...",
    5: "삭제된 Q-Type QN 항목을 분석 중입니다...",
    6: "각 QN의 대응 기한과 현재 진행 상태(D3/D5/D8)를 확인 중입니다...",
    7: "KPI 점수를 계산 중입니다...",
    8: "최신 정보로 보고서 메일 본문을 생성 중입니다...",
    9: "Outlook에서 보고서를 열고 있습니다. 확인 후 발송해주세요.",
}


def _main_impl(progress_cb=None, message_cb=None):
    """전체 파이프라인 실행.

    Args:
        progress_cb: 진행 콜백 (step_idx[, msg])
        message_cb:  상세 메시지 콜백 (step_idx[, msg])
    """
    def _safe_call(cb, step_idx, text):
        if not cb:
            return
        try:
            cb(step_idx, text)
        except TypeError:
            try:
                cb(step_idx)
            except Exception as e:
                log(f"[CALLBACK_ERR] step={step_idx} cb={cb} err={e}")
        except Exception as e:
            log(f"[CALLBACK_ERR] step={step_idx} cb={cb} err={e}")

    def emit(step_idx, msg=None):
        text = msg or STEP_MESSAGES.get(step_idx, "")
        log(f"[EMIT] step={step_idx} msg={text}")
        _safe_call(progress_cb, step_idx, text)
        _safe_call(message_cb, step_idx, text)

    def _fmt_new(kind_label, rows, detail_map):
        if not rows:
            return f"새롭게 발생한 {kind_label} QN은 없습니다."

        parts = []
        for r in rows:
            qn = str(r.get("QN", "")).strip()
            d = detail_map.get(qn, {}) if detail_map else {}
            created = r.get("Created", "") or "-"
            qn_type = d.get("QN_Type", "") or "-"
            material = d.get("Material", "") or "-"
            issue = (d.get("Issue", "") or "-").replace("\n", " ").strip()
            parts.append(
                f"{qn} | 등록일 {created} | TYPE {qn_type} | 품번 {material} | 이슈 {issue}"
            )

        joined = "; ".join(parts)
        return f"기존 보고서에 없던 새로운 {kind_label} QN을 확인했습니다. 신규 QN# [{joined}]"

    def _fmt_deleted(kind_label, deleted_list):
        if not deleted_list:
            return f"삭제된 {kind_label} QN은 없습니다."
        return (
            f"기존 보고서에는 있었지만 오늘은 사라진 {kind_label} QN을 확인했습니다. "
            f"삭제 QN# [{', '.join(deleted_list)}]"
        )

    def _emit_kpi(total, sat_d3, sat_d5, sat_d8):
        if not total:
            emit(7, "KPI 계산 결과가 없습니다 (데이터 0건).")
            return

        msg = (
            f"총 건수: {total}\n"
            f"D-3C 만족: {sat_d3} / {total}\n"
            f"D-5C 만족: {sat_d5} / {total}\n"
            f"D-8 만족: {sat_d8} / {total}"
        )
        emit(7, msg)

    print("[DEBUG] main() 진입")

    def _activate_mail(mail_item):
        """해당 메일 인스펙터를 활성화해 Word 편집/스타일 대상이 올바르도록 보장."""
        try:
            insp = mail_item.GetInspector()
            insp.Display()
            try:
                insp.Activate()
            except Exception:
                pass
        except Exception:
            log_exc("ACTIVATE_MAIL")

    # QN_Automailer가 주입하지 않았을 때만 settings 파일을 적용
    global _SETTINGS_APPLIED

    def _is_already_injected() -> bool:
        if START_DATE or END_DATE:
            return True
        if MAIL_RECIPIENTS and MAIL_RECIPIENTS != DEFAULT_MAIL_RECIPIENTS:
            return True
        if TEST_DATE != DEFAULT_TEST_DATE:
            return True
        if MAIL_SUBJECT_KEY != DEFAULT_MAIL_SUBJECT_KEY:
            return True
        if USE_ACTIVE_MAIL_FIRST != DEFAULT_USE_ACTIVE_MAIL_FIRST:
            return True
        if CONFIG != _DEFAULT_CONFIG:
            return True
        return False

    if not _SETTINGS_APPLIED:
        if _is_already_injected():
            _SETTINGS_APPLIED = True
        else:
            try:
                import settings_manager
                settings = settings_manager.load_settings()
                # 날짜/메일 주입
                TEST_DATE_local = settings_manager.get_effective_execution_date(settings)
                START_local, END_local = settings_manager.get_effective_date_range(settings)
                MAILS_local = settings_manager.get_mail_recipients_list(settings)

                globals()["TEST_DATE"] = TEST_DATE_local
                globals()["START_DATE"] = START_local
                globals()["END_DATE"] = END_local
                globals()["MAIL_RECIPIENTS"] = MAILS_local or list(DEFAULT_MAIL_RECIPIENTS)

                # 부가 설정 (메일/활성메일/제목 키)
                globals()["USE_ACTIVE_MAIL_FIRST"] = settings.get("use_active_mail_first", USE_ACTIVE_MAIL_FIRST)
                globals()["MAIL_SUBJECT_KEY"] = settings.get("mail_subject_key", MAIL_SUBJECT_KEY)

                # CONFIG 반영
                cfg = CONFIG.copy()
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
                CONFIG.update(cfg)
                _SETTINGS_APPLIED = True
                print("[DEBUG] settings_manager 적용 완료")
            except Exception:
                log_exc("settings_manager 적용 실패")

    day = TEST_DATE
    start_date = START_DATE or dt.date(day.year, 1, 1)
    end_date = END_DATE or day
    EXPORT_DIR = CONFIG["EXPORT_DIR"]

    log(
        f"[EFFECTIVE_SETTINGS] TEST_DATE={day}, RANGE=({start_date}~{end_date}), "
        f"EXPORT_DIR={EXPORT_DIR}, RECIPIENTS={MAIL_RECIPIENTS}"
    )

    # 필수 경로 검증
    if not EXPORT_DIR or not os.path.isdir(EXPORT_DIR):
        raise RuntimeError(f"EXPORT_DIR 경로가 없거나 잘못되었습니다: {EXPORT_DIR}")

    emit(0, f"SAP 트랜잭션 준비 중... (기간: {start_date}~{end_date})")

    # =========================
    # 🔍 가장 최근 저장된 파일 탐색 함수
    # =========================
    def find_latest_existing_file(base_dir, file_type="first", before_date=None):
        """
        지정 폴더에서 가장 최근 저장된 파일 경로를 반환
        - file_type: "first" → YYYYMMDD.txt 형식
                      "second" → YYYYMMDD_Q1.txt 형식
        - before_date: 오늘 날짜 이전 파일만 검색
        """
        pattern = r"(\d{8})(?:_Q1)?\.txt$"  # YYYYMMDD or YYYYMMDD_Q1
        all_files = os.listdir(base_dir)
        candidates = []
        for fn in all_files:
            m = re.match(pattern, fn)
            if not m:
                continue

            # 파일명 규칙 필터링
            if file_type == "first" and "_Q1" in fn:
                continue
            if file_type == "second" and "_Q1" not in fn:
                continue

            try:
                fdate = dt.datetime.strptime(m.group(1), "%Y%m%d").date()
            except:
                continue

            if before_date and fdate >= before_date:
                continue
            candidates.append((fdate, fn))

        if not candidates:
            return None

        candidates.sort(key=lambda x: x[0], reverse=True)
        latest_file = os.path.join(base_dir, candidates[0][1])
        return latest_file

    # =========================
    # 0) SAP 접속 → 지정된 기간 기준 파일 보장
    # =========================
    ensure_files_for_test_date(end_date, start_date=start_date)
    emit(0, "오늘 SAP에서 전체 QN 데이터를 불러왔습니다.")

    # =========================
    # 🔍 비교용 파일 자동 탐색
    # =========================
    latest_first_path  = find_latest_existing_file(EXPORT_DIR, "first", before_date=day)
    latest_second_path = find_latest_existing_file(EXPORT_DIR, "second", before_date=day)

    f_first_prev  = latest_first_path  if latest_first_path  else None
    f_second_prev = latest_second_path if latest_second_path else None

    log(f"[INFO] 최근 저장 파일: 표1={latest_first_path} / 표2={latest_second_path}")

    # =========================
    # 1) 메일 원본 확보
    # =========================
    src = None
    if USE_ACTIVE_MAIL_FIRST:
        src = get_active_mail()
        if src:
            log(f"[MAIL] 활성 메일 사용: {src.Subject}")
    
    if not src:
        src = find_latest_qn_mail_before_date(day)
        if not src:
            log("[MAIL] QN 관련 메일이 존재하지 않음 → 초기 설정 모드 실행")
            emit(1, "기존 보고서를 찾지 못해 새 보고서를 생성합니다.")
            src = build_first_mail_from_files(day)
            log("[MAIL] 초기 보고서 생성 후 파이프라인을 종료합니다 (중복 실행 방지)")
            emit(9, "Outlook에서 초기 보고서가 열렸습니다. 확인 후 발송해주세요.")
            return
        else:
            emit(1, "기존 보고서를 불러왔으며 업데이트 모드로 진행합니다.")
    else:
        emit(1, "기존 보고서를 불러왔으며 업데이트 모드로 진행합니다.")

    emit(1)

    src_date = (src.SentOn.date() if hasattr(src, "SentOn") and src.SentOn else day)
    html = src.HTMLBody or ""
    tables = get_tables(html)
    first_idx, first_table_html, second_idx, second_table_html = detect_anchor_tables(tables)

    # =========================
    # 2) 첫/두 번째 표에서 앵커 QN 후보 추출
    # =========================
    first_candidates = collect_qn_candidates_from_table(first_table_html, limit=30)
    second_candidates = collect_qn_candidates_from_table(second_table_html, limit=30)
    last_qn_first = first_candidates[0] if first_candidates else None
    last_qn_second = second_candidates[0] if second_candidates else None
    log(
        f"[ANCHOR] First(idx={first_idx}) last={last_qn_first} 후보={first_candidates[:5]} / "
        f"Second(idx={second_idx}) last={last_qn_second} 후보={second_candidates[:5]}"
    )

    # =========================
    # 3) 오늘자 파일 경로
    # =========================
    f_first_day   = path_first(day)
    f_second_day  = path_second(day)
    f_znot_day    = path_znot(day)

    # =========================
    # 4) 신규/삭제 계산
    # =========================
    new_rows_first  = parse_export_txt(
        f_first_day,
        last_qn_first,
        anchor_candidates=first_candidates[1:] if len(first_candidates) > 1 else None,
    )

    new_rows_second = parse_export_txt(
        f_second_day,
        last_qn_second,
        anchor_candidates=second_candidates[1:] if len(second_candidates) > 1 else None,
    )
    log(f"[NEW] 첫표 신규={len(new_rows_first)} (예: {new_rows_first[:5]})")
    log(f"[NEW] 두표 신규={len(new_rows_second)} (예: {new_rows_second[:5]})")

    # =========================
    # 5) 신규 QN 상세 조회(QM03)
    # =========================
    detail_map_first  = get_details_for_qns_first([r["QN"] for r in new_rows_first])   if new_rows_first  else {}
    detail_map_second = get_details_for_qns_second([r["QN"] for r in new_rows_second]) if new_rows_second else {}
    detail_map = {**detail_map_first, **detail_map_second}

    # 신규/삭제 결과 메시지 출력 (상세 포함)
    emit(2, _fmt_new("C-Type", new_rows_first, detail_map_first))
    emit(4, _fmt_new("Q-Type", new_rows_second, detail_map_second))

    deleted_first = []
    deleted_second = []

    if f_first_prev and os.path.isfile(f_first_prev):
        deleted_first = detect_deleted(f_first_prev, f_first_day)
    else:
        log("[WARN] 첫 번째 표 비교 대상 파일 없음 → 삭제 QN 계산 생략")

    emit(3, _fmt_deleted("C-Type", deleted_first))

    if f_second_prev and os.path.isfile(f_second_prev):
        deleted_second = detect_deleted(f_second_prev, f_second_day)
    else:
        log("[WARN] 두 번째 표 비교 대상 파일 없음 → 삭제 QN 계산 생략")

    emit(5, _fmt_deleted("Q-Type", deleted_second))

    # =========================
    # 6) ZNOT 보강
    # =========================
    def blank_qns_first(table_html):
        rows = table_rows(table_html)
        if len(rows) < 2:
            return []
        hdrs = header_cells(rows[0])
        dcols = []
        for i, h in enumerate(hdrs):
            H = h.upper().replace(" ", "")
            if any(tag in H for tag in ("D-3C", "D-5C", "D-8")):
                dcols.append(i)
        qn_idx = find_qn_index_from_header(hdrs)
        if qn_idx is None:
            return []
        blanks = []
        for r in rows[1:]:
            cells = data_cells(r)
            if qn_idx >= len(cells):
                continue
            qn = cells[qn_idx].strip()
            if not qn:
                continue
            if any(ci < len(cells) and cells[ci] == "" for ci in dcols):
                blanks.append(qn)
        return blanks

    blanks = blank_qns_first(first_table_html)
    today_first_qns = load_qn_list(f_first_day)
    target_for_znot = set(today_first_qns)
    if not target_for_znot:
        target_for_znot = set(blanks)
    else:
        target_for_znot.update(blanks)
    target_for_znot.update(r.get("QN") for r in new_rows_first if r.get("QN"))

    znot_rows = extract_d_values_from_znot(f_znot_day, list(target_for_znot)) if target_for_znot else []
    zmap = {}
    for r in znot_rows:
        zmap.setdefault(r["QN"], []).append({"Task": r["Task"], "Elapsed": r["Elapsed"]})
    emit(6, "각 QN의 대응 기한과 현재 진행 상태를 확인했습니다.")

    # =========================
    # 7) 메일 복제 및 표 업데이트
    # =========================
    outlook, mapi = _get_outlook_namespace()

    # --- Outbox → Drafts 이동 ---
    outbox = mapi.GetDefaultFolder(4)
    drafts = mapi.GetDefaultFolder(16)
    outbox_items = outbox.Items
    move_count = 0
    for i in range(outbox_items.Count, 0, -1):
        item = outbox_items.Item(i)
        if getattr(item, "Class", None) != 43:
            continue
        subj = getattr(item, "Subject", "")
        if MAIL_SUBJECT_KEY in subj:
            try:
                item.Move(drafts)
                move_count += 1
            except Exception as e:
                log(f"[OUTBOX→DRAFTS 이동 실패] {subj}: {e}")
    log(f"[CLEANUP] Outbox → Drafts 이동 완료: {move_count}건")

    # --- Drafts 내 동일 제목 메일 삭제 ---
    drafts_items = drafts.Items
    del_count = 0
    for i in range(drafts_items.Count, 0, -1):
        item = drafts_items.Item(i)
        if getattr(item, "Class", None) != 43:
            continue
        subj = getattr(item, "Subject", "")
        if MAIL_SUBJECT_KEY in subj:
            try:
                item.Delete()
                del_count += 1
            except Exception as e:
                log(f"[DRAFTS 삭제 실패] {subj}: {e}")
    log(f"[CLEANUP] Drafts 내 '{MAIL_SUBJECT_KEY}' 삭제 완료: {del_count}건")

    # --- 새 메일 생성 ---
    new_mail = outlook.CreateItem(0)
    try:
        new_mail.BodyFormat = 2  # olFormatHTML
    except Exception:
        pass
    new_subject = replace_date_in_subject(src.Subject, src_date, day)
    new_html = replace_date_in_body(html, src_date, day)

    new_mail.Subject = new_subject
    new_mail.HTMLBody = new_html
    recipients = MAIL_RECIPIENTS if MAIL_RECIPIENTS else list(DEFAULT_MAIL_RECIPIENTS)
    new_mail.To = "; ".join(recipients)
    new_mail.Display()
    _activate_mail(new_mail)
    log("[MAIL] 새 메일 창 표시(원본 스타일 복제 + 날짜 치환)")

    # =========================
    # 8) 표 업데이트 (삭제 + 신규 + 색상)
    # =========================
    cur_html = new_mail.HTMLBody or ""
    cur_tables = get_tables(cur_html)
    _cur_first_idx, t1, _cur_second_idx, t2 = detect_anchor_tables(cur_tables)

    # --- 첫 번째 표 ---
    t1_updated, removed1 = remove_deleted_from_table(t1, deleted_first)
    header_map_first = {
        "Created": "등록일",
        "QN": "QN #",
        "QN_Type": "QN TYPE",
        "Material": "품번",
        "Issue": ["이슈 내용", "이슈", "이슈내용"],
        "D3C": "D-3C",
        "D5C": "D-5C",
        "D8": "D-8",
    }
    new_rows_for_mail_1 = []
    for r in new_rows_first:
        qn = r.get("QN", "")
        d  = detail_map.get(qn, {})
        z  = zmap.get(qn, [])
    
        # ZNOT에서 D-3C/5C/8 찾아서 값 매핑
        d3 = next((x["Elapsed"] for x in z if (x.get("Task") or "").upper().replace(" ", "") == "D-3C"), "")
        d5 = next((x["Elapsed"] for x in z if (x.get("Task") or "").upper().replace(" ", "") == "D-5C"), "")
        d8 = next((x["Elapsed"] for x in z if (x.get("Task") or "").upper().replace(" ", "") == "D-8"),  "")
    
        new_rows_for_mail_1.append({
            "QN": qn,
            "QN_Type": d.get("QN_Type", ""),
            "Material": d.get("Material", ""),
            "Issue": d.get("Issue", ""),
            "Created": r.get("Created", ""),
            "D3C": d3, "D5C": d5, "D8": d8,  # ✅ 여기가 핵심
        })

    t1_updated = append_new_rows_first(t1_updated, new_rows_for_mail_1, header_map_first)
     # ✅ 기존행 빈칸 보강
    t1_updated = fill_znot_into_first_table(t1_updated, zmap)

    # --- 두 번째 표 (Outlook 기반 재조립) ---
    
    # (1) 기존 표 파싱
    rows = table_rows(t2)
    if len(rows) < 2:
        log("[MAIL/SECOND] 기존 표 행이 부족합니다.")
        parsed_existing = []
        hdrs = ["No.", "등록일", "QN #", "QN TYPE", "품번", "이슈 내용"]
    else:
        hdrs = [strip_html(c) for c in re.findall(r"(?is)<t[dh][^>]*>.*?</t[dh]>", rows[0])]
        # ✅ 공백 헤더(&nbsp; 등) 제거
        hdrs = [h for h in hdrs if h.strip() not in ("", "&nbsp;")]
        hdr_cnt = len(hdrs)
        parsed_existing = []
        # ✅ start=2 → 제거 (Outlook이 자동 생성한 빈 <tr> 2개 무시)
        for row_html in rows[1:]:
            cells = [strip_html(c) for c in re.findall(r"(?is)<t[dh][^>]*>.*?</t[dh]>", row_html)]
            if not any(c.strip() for c in cells) or all(c.replace("&nbsp;", "").strip() == "" for c in cells):
                continue
            if len(cells) < hdr_cnt:
                cells += [""] * (hdr_cnt - len(cells))
            elif len(cells) > hdr_cnt:
                cells = cells[:hdr_cnt]
            parsed_existing.append(dict(zip(hdrs, cells)))
        log(f"[MAIL/SECOND] 기존 표 파싱 완료: {len(parsed_existing)}건 (열={hdr_cnt})")
    
    # (2) 삭제 반영
    if deleted_second:
        before = len(parsed_existing)
        parsed_existing = [
            r for r in parsed_existing
            if (r.get("QN #", "").strip() or r.get("QN", "").strip()) not in deleted_second
        ]
        log(f"[MAIL/SECOND] 삭제 반영 완료: {before} → {len(parsed_existing)}건")
    

    # (3) 신규 추가
    for r in new_rows_second:
        qn = str(r.get("QN", "")).strip()
        if not qn:
            continue
    
        # 기존 행 중 이미 같은 QN이 있으면 skip
        exists = any(
            (row.get("QN #", "").strip() or row.get("QN", "").strip()) == qn
            for row in parsed_existing
        )
        if exists:
            continue
    
        d = detail_map.get(qn, {})
    
        # ✅ 표준 키 기반 dict 생성
        new_row = {
            "No": "",
            "Created": r.get("Created", ""),
            "QN": qn,
            "QN_Type": d.get("QN_Type", ""),
            "Material": d.get("Material", ""),
            "Issue": d.get("Issue", ""),
        }
    
        parsed_existing.append(new_row)

    log(f"[MAIL/SECOND] 신규 행 추가 완료 후 총 {len(parsed_existing)}건")
    
    # (4) 번호 재정렬
    for i, row in enumerate(parsed_existing, start=1):
        row["No."] = str(i)

    hdrs = [re.sub(r"[\u200b\u200c\u200d\ufeff\xa0]", "", h).strip() for h in hdrs]
    # 헤더 → 표준필드 키 매핑 (인덱스 기반)
    # 표준필드 키: "No", "Created", "QN", "QN_Type", "Material", "Issue"
    selector = {}
    for i, h in enumerate(hdrs):
        H = _norm_header(h)
    
        # ✅ QN TYPE 먼저 검사 (QN보다 앞에!)
        if any(k in H for k in ["qntype", "type"]):
            selector[i] = "QN_Type"
        elif any(k in H for k in ["no", "번호"]):
            selector[i] = "No"
        elif any(k in H for k in ["등록일", "created", "creationdate"]):
            selector[i] = "Created"
        elif any(k in H for k in ["qn#", "notification", "note", "notenumber", "qn"]):
            selector[i] = "QN"
        elif any(k in H for k in ["품번", "material", "partno", "partnumber"]):
            selector[i] = "Material"
        elif any(k in H for k in ["이슈내용", "이슈", "issue", "description", "shorttext", "text"]):
            selector[i] = "Issue"

    # (5) 새 HTML 조립
    base_style = "font-family:'Malgun Gothic'; font-size:10pt; text-align:center;"
    rows_html = []

    # 5-1) 헤더 행
    header_html = "<tr>" + "".join(
        f"<th style='{base_style} font-weight:normal;'>{h}</th>" for h in hdrs
    ) + "</tr>"
    rows_html.append(header_html)

    # 5-2) 데이터 행
    for row in parsed_existing:
        tds = []
        for col_idx, h in enumerate(hdrs):
            # 1️⃣ 기존행(헤더키) 우선
            val = row.get(h, None)
            # 2️⃣ 없으면 매핑된 표준키로 재조회
            if val in (None, ""):
                std_key = selector.get(col_idx)
                if std_key:
                    val = row.get(std_key, "")
            tds.append(f"<td style='{base_style}'>{sanitize_cell_value(val)}</td>")
        rows_html.append("<tr>" + "".join(tds) + "</tr>")

    # 최종 완성
    t2_final = (
        "<table style='border-collapse:collapse; table-layout:fixed; width:auto; font-family:\"Malgun Gothic\"; font-size:10pt;'>"
        + "".join(rows_html)
        + "</table>"
    )
    # ✅ 두 번째 표 안정화(셀수/공백 보정)
    t2_final = sanitize_table_second(t2_final)

    


    # --- 번호 재정렬 ---
    t1_final = renumber_first_table(t1_updated)
    
    # 첫 번째 표 값만 정리 (구조 유지)
    t1_final = sanitize_table_cells_preserve_structure(t1_final)
    
    
    # --- 메일 본문 갱신 ---
    new_body = cur_html.replace(t1, t1_final, 1).replace(t2, t2_final, 1)
    new_body = finalize_outlook_html(new_body)
    new_mail.HTMLBody = new_body

    # --- 색상/정렬 적용 ---
    _activate_mail(new_mail)
    apply_final_table_styles_then_colors(
        data_start_row=3,
        apply_header_style=True,
        clear_header_emphasis=True,
    )

    # ✅ 요약 블록 삽입 (총 건수 및 D-3/5/8 비율)
    total = sat_d3 = sat_d5 = sat_d8 = 0
    try:
        _activate_mail(new_mail)
        mail_obj, total, sat_d3, sat_d5, sat_d8 = parse_first_table_C1Ck_2row_header(new_mail)
        append_summary_in_mail(mail_obj, total, sat_d3, sat_d5, sat_d8)
        log("[MAIL] 요약 블록 추가 완료")
    except Exception:
        log_exc("SUMMARY_APPEND")

    try:
        apply_uniform_body_font(new_mail)
    except Exception:
        log_exc("UNIFORM_FONT")

    _emit_kpi(total, sat_d3, sat_d5, sat_d8)
    emit(8, "최신 정보로 보고서 메일 본문을 생성했습니다.")
    emit(9, "Outlook에서 보고서가 열렸습니다. 확인 후 발송해주세요.")

    log("✅ 전체 파이프라인 완료 (모든 로그 참고)")


def main(progress_cb=None, message_cb=None):
    with sap_com_apartment():
        return _main_impl(progress_cb, message_cb)


if __name__ == "__main__":
    try:
        with sap_com_apartment():
            # 1️⃣ 메일 생성 / 갱신
            _main_impl()
            log("✅ 1단계: 메일 생성 및 표 갱신 완료")

            # 2️⃣ 통계 요약 블록 추가
            mail, total, sat_d3, sat_d5, sat_d8 = parse_first_table_C1Ck_2row_header()
            append_summary_in_mail(mail, total, sat_d3, sat_d5, sat_d8)
            log("✅ 2단계: 통계 요약 블록 추가 완료")

    except Exception:
        log_exc("MAIN+SUMMARY")
