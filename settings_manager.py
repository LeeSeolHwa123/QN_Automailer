"""
설정 관리 모듈 (JSON 저장/불러오기 + 유효성 검증)
"""
import json
import os
import re
from datetime import datetime as _dt, date as _date
from pathlib import Path

DEFAULT_SETTINGS_FILE = "qn_automailer_settings.json"

# 기본값 (설정 없을 때)
DEFAULT_SETTINGS = {
    "execution_date": None,  # None = today
    "date_range_start": None,  # None = 1/1 of current year
    "date_range_end": None,  # None = today
    "export_dir": "",  # 필수 (빈 값 = 미설정)
    "mail_recipients": "",  # 필수 (빈 값 = 미설정)
    "plant": "2201",
    "mail_subject_key": "QN 등록 현황 공유",
    "use_active_mail_first": True,
    "file_templates": {
        "first": "{YYYYMMDD}.txt",
        "second": "{YYYYMMDD}_Q1.txt",
        "znot": "{YYYYMMDD}_ZNOT.txt"
    },
    "types_first": ["C1", "C3", "C5", "C7", "G1", "G3", "G5", "G7", "N1", "CK"],
    "types_second": [f"Q{i}" for i in range(1, 10)],
    "auto_generate_files": True,
    "max_email_scan": 200,
    "cleanup_outbox_drafts": True,
    "apply_styles": True,
    "font_name": "맑은 고딕",
    "font_size": 10,
    "color_thresholds": {
        "d3": 1,
        "d5": 5,
        "d8": 20
    }
}


def validate_date_string(date_str: str) -> _date:
    """
    YYYYMMDD 형식 날짜 문자열을 date 객체로 변환 및 검증
    Returns: datetime.date 객체
    Raises: ValueError if invalid format or date
    """
    if not date_str or not isinstance(date_str, str):
        raise ValueError(f"Invalid date format: {date_str}")
    
    date_str = date_str.strip()
    if len(date_str) != 8 or not date_str.isdigit():
        raise ValueError(f"Date must be YYYYMMDD format, got: {date_str}")
    
    try:
        return _dt.strptime(date_str, "%Y%m%d").date()
    except ValueError as e:
        raise ValueError(f"Invalid date: {date_str} ({e})")


def validate_email_list(email_str: str) -> list:
    """
    세미콜론(;)으로 구분된 이메일 리스트 검증
    Returns: 검증된 이메일 리스트
    Raises: ValueError if invalid email format
    """
    if not email_str or not isinstance(email_str, str):
        raise ValueError("Recipients cannot be empty")
    
    emails = [e.strip() for e in email_str.split(";") if e.strip()]
    
    if not emails:
        raise ValueError("No valid email addresses found")
    
    # 간단한 이메일 포맷 검증
    email_pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    for email in emails:
        if not re.match(email_pattern, email):
            raise ValueError(f"Invalid email format: {email}")
    
    return emails


def validate_directory(dir_path: str) -> str:
    """
    디렉토리 경로 검증 (존재 확인)
    Returns: 절대경로
    Raises: ValueError if directory doesn't exist
    """
    if not dir_path or not isinstance(dir_path, str):
        raise ValueError("Export directory cannot be empty")
    
    dir_path = dir_path.strip()
    abs_path = str(Path(dir_path).resolve())
    
    if not os.path.isdir(abs_path):
        raise ValueError(f"Directory does not exist: {abs_path}")
    
    return abs_path


def load_settings(filepath: str = DEFAULT_SETTINGS_FILE) -> dict:
    """
    JSON 파일에서 설정 불러오기
    파일 없으면 기본값 반환
    Returns: settings dict
    """
    if not os.path.isfile(filepath):
        print(f"[Settings] No config file found ({filepath}), using defaults")
        return DEFAULT_SETTINGS.copy()
    
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            settings = json.load(f)
        
        # 기본값과 병합 (새로 추가된 항목 대비)
        merged = DEFAULT_SETTINGS.copy()
        merged.update(settings)
        
        print(f"[Settings] Loaded from {filepath}")
        return merged
    
    except Exception as e:
        print(f"[Settings] Error loading config ({e}), using defaults")
        return DEFAULT_SETTINGS.copy()


def save_settings(settings: dict, filepath: str = DEFAULT_SETTINGS_FILE) -> bool:
    """
    설정을 JSON 파일로 저장
    Returns: True if successful, False otherwise
    """
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False, default=str)
        print(f"[Settings] Saved to {filepath}")
        return True
    except Exception as e:
        print(f"[Settings] Error saving config ({e})")
        return False


def validate_settings(settings: dict) -> tuple[bool, str]:
    """
    사용자 입력 설정 유효성 검증 (필수값 + 포맷)
    Returns: (is_valid: bool, error_message: str)
    """
    errors = []
    
    # 필수값 검증
    if not settings.get("export_dir", "").strip():
        errors.append("Export directory is required")
    else:
        try:
            validate_directory(settings["export_dir"])
        except ValueError as e:
            errors.append(str(e))
    
    if not settings.get("mail_recipients", "").strip():
        errors.append("Mail recipients is required")
    else:
        try:
            validate_email_list(settings["mail_recipients"])
        except ValueError as e:
            errors.append(str(e))
    
    # 선택값 검증 (미입력시 기본값으로)
    if settings.get("execution_date"):
        try:
            validate_date_string(settings["execution_date"])
        except ValueError as e:
            errors.append(f"Invalid execution date: {e}")
    
    if settings.get("date_range_start"):
        try:
            validate_date_string(settings["date_range_start"])
        except ValueError as e:
            errors.append(f"Invalid date range start: {e}")
    
    if settings.get("date_range_end"):
        try:
            validate_date_string(settings["date_range_end"])
        except ValueError as e:
            errors.append(f"Invalid date range end: {e}")
    
    # 식물 코드 (선택: 길이 검증)
    if settings.get("plant"):
        plant = str(settings["plant"]).strip()
        if not (1 <= len(plant) <= 10):
            errors.append("Plant code must be 1-10 characters")
    
    # 숫자값 검증
    if settings.get("max_email_scan"):
        try:
            max_scan = int(settings["max_email_scan"])
            if max_scan < 1:
                errors.append("Max email scan must be >= 1")
        except ValueError:
            errors.append("Max email scan must be a number")
    
    # 색상 임계값
    for key in ["d3", "d5", "d8"]:
        if settings.get("color_thresholds", {}).get(key):
            try:
                val = int(settings["color_thresholds"][key])
                if val < 0:
                    errors.append(f"Threshold {key} must be >= 0")
            except ValueError:
                errors.append(f"Threshold {key} must be a number")
    
    if errors:
        return False, "\n".join(errors)
    
    return True, ""


def get_effective_execution_date(settings: dict) -> _date:
    """
    설정에서 실행 날짜 추출 (미입력시 today)
    """
    if settings.get("execution_date"):
        try:
            return validate_date_string(settings["execution_date"])
        except ValueError:
            pass
    return _date.today()


def get_effective_date_range(settings: dict) -> tuple[_date, _date]:
    """
    설정에서 날짜 조회 범위 추출
    미입력시: (당해 1월 1일, 오늘)
    """
    today = _date.today()
    
    start = today
    if settings.get("date_range_start"):
        try:
            start = validate_date_string(settings["date_range_start"])
        except ValueError:
            start = _date(today.year, 1, 1)
    else:
        start = _date(today.year, 1, 1)
    
    end = today
    if settings.get("date_range_end"):
        try:
            end = validate_date_string(settings["date_range_end"])
        except ValueError:
            end = today
    elif settings.get("execution_date"):
        try:
            end = validate_date_string(settings["execution_date"])
        except ValueError:
            end = today
    
    # 논리 검증: 시작일 <= 종료일
    if start > end:
        start, end = end, start
    
    return start, end


def get_mail_recipients_list(settings: dict) -> list:
    """
    설정에서 메일 수신자 리스트 추출
    Returns: 검증된 이메일 리스트
    """
    try:
        return validate_email_list(settings.get("mail_recipients", ""))
    except ValueError:
        return []
