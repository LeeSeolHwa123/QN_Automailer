# -*- coding: utf-8 -*-
"""
SAP 세션 공유 락 관리
여러 스레드에서 동시에 SAP 세션에 접근할 때 충돌 방지
"""
import threading
import time

# 전역 SAP 세션 접근 락
_SAP_LOCK = threading.RLock()  # 재진입 가능 락

class SAPLockContext:
    """
    SAP 세션 접근 보호 컨텍스트 매니저
    
    사용법:
        from sap_shared_lock import SAPLockContext
        
        with SAPLockContext():
            # SAP 세션 접근 코드
            session = get_session()
            session.findById("...").text = "..."
    """
    
    def __init__(self, timeout=30):
        """
        Args:
            timeout: 락 획득 대기 시간 (초)
        """
        self.timeout = timeout
        self.acquired = False
    
    def __enter__(self):
        """컨텍스트 진입 - 락 획득"""
        start_time = time.time()
        while True:
            if _SAP_LOCK.acquire(blocking=False):
                self.acquired = True
                return self
            
            if time.time() - start_time > self.timeout:
                raise TimeoutError(f"SAP 락 획득 실패 (타임아웃: {self.timeout}초)")
            
            time.sleep(0.1)
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """컨텍스트 종료 - 락 해제"""
        if self.acquired:
            _SAP_LOCK.release()
            self.acquired = False
        return False


def acquire_sap_lock():
    """수동으로 SAP 락 획득"""
    return _SAP_LOCK.acquire(blocking=True, timeout=30)


def release_sap_lock():
    """수동으로 SAP 락 해제"""
    try:
        _SAP_LOCK.release()
    except RuntimeError:
        pass  # 이미 해제된 경우 무시
