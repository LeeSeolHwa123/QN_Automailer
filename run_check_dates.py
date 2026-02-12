import traceback
from datetime import date
import qn_full_pipeline as fp


def cb(step, msg=None):
    print(f"[PROG]{step}: {msg}")


def run_for(d):
    print("\n--- RUN FOR:", d)
    try:
        fp.TEST_DATE = d
        # ensure no leftover injected START/END
        fp.START_DATE = None
        fp.END_DATE = None
        fp.MAIL_SUBJECT_KEY = fp.MAIL_SUBJECT_KEY
        fp.USE_ACTIVE_MAIL_FIRST = fp.USE_ACTIVE_MAIL_FIRST
        fp.MAIL_RECIPIENTS = fp.MAIL_RECIPIENTS
        # call main (which wraps sap_com_apartment)
        fp.main(progress_cb=cb, message_cb=cb)
    except Exception:
        traceback.print_exc()

if __name__ == '__main__':
    run_for(date(2026,2,10))
    run_for(date(2026,2,11))
