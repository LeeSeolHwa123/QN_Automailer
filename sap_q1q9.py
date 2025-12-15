# -*- coding: utf-8 -*-
"""
Q1-Q9 Notification Type 조회 (두 번째 표)
QN_Automailer.py의 multisel_types 로직과 일치하도록 수정됨
"""
import time
import datetime as dt
import win32com.client
import pythoncom
from sap_shared_lock import SAPLockContext

def get_session():
    """SAP 세션 획득 - 최대 재시도"""
    max_retries = 5
    for attempt in range(max_retries):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            print(f"[Q1Q9] Connections: {application.Children.Count}")
            for i in range(application.Children.Count):
                conn = application.Children(i)
                if not conn.DisabledByServer and conn.Children.Count > 0:
                    session = conn.Children(0)
                    try:
                        _ = session.findById("wnd[0]")
                        print(f"[Q1Q9] Session found at attempt {attempt}")
                        return session
                    except:
                        pass
            if attempt < max_retries - 1:
                wait_time = 2.0 + attempt
                print(f"[Q1Q9] No session found, retry {attempt + 1}/{max_retries} in {wait_time}s")
                time.sleep(wait_time)
        except Exception as e:
            if attempt < max_retries - 1:
                wait_time = 2.0 + attempt
                print(f"[Q1Q9] Error: {e}, retry {attempt + 1}/{max_retries} in {wait_time}s")
                time.sleep(wait_time)
    raise RuntimeError("No active SAP session after 5 retries")

def wait_ready(session, timeout=12):
    """SAP Busy 해제 대기"""
    app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if not app.Busy and not session.Busy:
                session.findById("wnd[0]")
                return True
        except:
            pass
        time.sleep(0.15)
    return False

def wait_for_obj(session, obj_id, timeout=8, interval=0.2):
    t0 = time.time()
    last_err = None
    while time.time() - t0 < timeout:
        try:
            return session.findById(obj_id)
        except Exception as e:
            last_err = e
            time.sleep(interval)
    raise TimeoutError(f"객체 대기 타임아웃: {obj_id} / {last_err}")

def safe_settext(session, obj_id, text):
    ctrl = wait_for_obj(session, obj_id, timeout=5)
    ctrl.text = str(text)
    try:
        ctrl.setFocus()
    except:
        pass

def safe_press(session, obj_id):
    ctrl = wait_for_obj(session, obj_id, timeout=5)
    ctrl.press()

def go_tcode(session, tcode):
    ok = wait_for_obj(session, "wnd[0]/tbar[0]/okcd", timeout=5)
    ok.text = f"/n{tcode}"
    session.findById("wnd[0]").sendVKey(0)
    wait_ready(session, timeout=8)

def close_popup_if_any(session):
    """열린 모달(wnd[1])이 있으면 Enter로 닫기"""
    try:
        if session.FindById("wnd[1]", False):
            session.findById("wnd[1]").sendVKey(0)
            wait_ready(session, timeout=5)
            return True
    except:
        pass
    return False

def set_search_filters(session):
    """검색조건 입력: 기간, Plant, 상태"""
    TODAY = dt.date.today()
    YEAR_START = dt.date(TODAY.year, 1, 1)
    PLANT = "2201"
    
    print("[Q1Q9] Setting search filters...")
    safe_settext(session, "wnd[0]/usr/ctxt[4]", YEAR_START.strftime("%Y/%m/%d"))
    safe_settext(session, "wnd[0]/usr/ctxt[5]", TODAY.strftime("%Y/%m/%d"))
    safe_settext(session, "wnd[0]/usr/ctxt[8]", PLANT)

    wait_for_obj(session, "wnd[0]/usr/chk[0]").selected = True      # Open
    wait_for_obj(session, "wnd[0]/usr/chk[2]").selected = True      # Being processed
    wait_for_obj(session, "wnd[0]/usr/chk[3]").selected = True      # Completed
    wait_for_obj(session, "wnd[0]/usr/chk[1]").selected = True      # Postponed

    print("[Q1Q9] Search filters set OK")

def set_notif_types_via_popup(session, notif_list):
    """
    Notification Type 다중 입력 (Q1-Q9)
    QN_Automailer.py의 multisel_types 로직과 동일
    """
    print("[Q1Q9] Opening multi-select popup...")
    safe_press(session, "wnd[0]/usr/btn[1]")  # Multiple Selection 버튼

    TABLE_BASE = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssub/1/2/tblSAPLALDBSINGLE"
    wait_for_obj(session, TABLE_BASE, timeout=8)

    # 첫 페이지: 첫 8개 입력
    first_page_count = min(8, len(notif_list))
    for idx in range(first_page_count):
        cell_id = f"{TABLE_BASE}/ctxt[1,{idx}]"
        safe_settext(session, cell_id, notif_list[idx])
        time.sleep(0.25)

    print(f"[Q1Q9] First page input OK: {notif_list[:first_page_count]}")

    # 커밋 (Enter)
    if first_page_count > 0:
        try:
            last_cell = f"{TABLE_BASE}/ctxt[1,{first_page_count-1}]"
            session.findById(last_cell).setFocus()
            session.findById("wnd[1]").sendVKey(0)
        except Exception as e:
            print(f"[Q1Q9] First page commit failed: {e}")

    wait_ready(session, timeout=4)
    time.sleep(0.4)

    # 남은 타입이 있으면 PageDown 후 입력
    if len(notif_list) > 8:
        try:
            print("[Q1Q9] PageDown to next page...")
            session.findById("wnd[1]").sendVKey(82)
            wait_ready(session, timeout=6)
            time.sleep(1.0)
        except Exception as e:
            print(f"[Q1Q9] PageDown failed: {e}")

        last_first_page_value = notif_list[7]
        remaining = [last_first_page_value] + notif_list[8:]

        print(f"[Q1Q9] Second page input: {remaining}")

        for idx, val in enumerate(remaining):
            try:
                cell_id = f"{TABLE_BASE}/ctxt[1,{idx}]"
                safe_settext(session, cell_id, val)
                time.sleep(0.3)
            except Exception as e:
                print(f"[Q1Q9] Second page input failed for {val}: {e}")

        # 커밋
        try:
            last_idx = len(remaining) - 1
            last_cell = f"{TABLE_BASE}/ctxt[1,{last_idx}]"
            session.findById(last_cell).setFocus()
            session.findById("wnd[1]").sendVKey(0)
        except Exception as e:
            print(f"[Q1Q9] Second page commit failed: {e}")

        wait_ready(session, timeout=4)
        time.sleep(0.4)

    # 팝업 닫기
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

    print(f"[Q1Q9] Multi-select popup closed, OK: {closed}")

def execute_search(session):
    """실행 (=ONLI)"""
    print("[Q1Q9] Executing search...")
    close_popup_if_any(session)
    ok = wait_for_obj(session, "wnd[0]/tbar[0]/okcd", timeout=4)
    ok.text = "=ONLI"
    session.findById("wnd[0]").sendVKey(0)
    wait_ready(session, timeout=10)
    print("[Q1Q9] Search complete")

def close_any_popups(session):
    """모든 열린 팝업 닫기"""
    for _ in range(5):
        try:
            if session.FindById("wnd[1]", False):
                session.findById("wnd[1]").sendVKey(0)
                time.sleep(0.3)
        except:
            break

def run_q1q9_search():
    """Q1-Q9 조회 메인 플로우"""
    from sap_shared_lock import SAPLockContext
    
    with SAPLockContext():  # 공유 스레드 락으로 SAP 세션 보호
        try:
            session = get_session()
            print("[OK] Connected to SAP")
            
            # 이전 작업 팝업 정리
            close_any_popups(session)
            time.sleep(1.0)

            try:
                go_tcode(session, "ZQCN")
                time.sleep(0.5)
            except Exception as e:
                print("[WARN] TCode move skipped or failed:", e)

            NOTIF_TYPES = ["Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Q9"]
            
            set_search_filters(session)
            set_notif_types_via_popup(session, NOTIF_TYPES)
            execute_search(session)
        except Exception as e:
            print(f"[ERROR] Q1-Q9 search failed: {e}")
            raise

def main():
    """메인 진입점"""
    pythoncom.CoInitialize()
    try:
        run_q1q9_search()
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
