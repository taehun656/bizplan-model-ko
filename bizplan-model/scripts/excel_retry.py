"""Retry the two COM statuses that mean Excel temporarily refused the call."""
from contextlib import contextmanager
import time


BUSY_MESSAGE = (
    "Excel이 응답하지 않아 검증을 끝내지 못했습니다. "
    "모델의 오류가 아닐 수 있으니 Excel이 준비된 뒤 검증을 다시 실행하세요."
)


def busy(exc):
    while exc is not None:
        if getattr(exc, "hresult", None) in (-2147418111, -2147417846):
            return True
        exc = exc.__cause__
    return False


def retry_excel(action, attempts=6):
    for attempt in range(1, attempts + 1):
        try:
            return action()
        except Exception as exc:
            if not busy(exc) or attempt == attempts:
                raise
            time.sleep(min(0.25 * attempt, 1.0))

class ExcelShutdownBusy(RuntimeError):
    """The private server kept refusing Quit; starting another one would compound the failure."""


def quit_excel(action, retry=retry_excel):
    try:
        retry(action)
    except Exception as exc:
        if busy(exc):
            raise ExcelShutdownBusy(BUSY_MESSAGE) from None
        raise


def cleanup_after_error(action, primary):
    """Prefer a genuine operation error when its cleanup also fails."""
    try:
        action()
    except Exception:
        if busy(primary):
            raise


@contextmanager
def cleanup_excel(action):
    try:
        yield
    except BaseException as primary:
        cleanup_after_error(action, primary)
        raise
    else:
        action()
