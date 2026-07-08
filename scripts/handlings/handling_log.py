import logging
import traceback
from logging.handlers import RotatingFileHandler

import pandas as pd

logger = logging.getLogger("ReportConversion")


def setup_logger(logfile: str = "report_conv.log"):
    logger.setLevel(logging.DEBUG)
    if logger.handlers:
        return logger

    fh = RotatingFileHandler(logfile, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(name)s: %(message)s')
    fh.setFormatter(formatter)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger





def attempt_recover(callable_func, recover_funcs=None, attempts: int = 2, *args, **kwargs):
    """Выполнить функцию с логированием ошибок и попытками восстановления.

    callable_func: функция, которую нужно выполнить
    recover_funcs: список функций без аргументов, которые попытаются восстановить состояние
    attempts: сколько раз попытаться выполнить (включая первую)
    """
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            return callable_func(*args, **kwargs)
        except Exception as e:
            last_exc = e
            logger.exception("Error on attempt %s running %s: %s", attempt, getattr(callable_func, "__name__", str(callable_func)), e)
            if recover_funcs:
                for rf in recover_funcs:
                    try:
                        logger.info("Running recovery function: %s", getattr(rf, "__name__", str(rf)))
                        rf()
                    except Exception:
                        logger.exception("Recovery function failed: %s", getattr(rf, "__name__", str(rf)))
            if attempt < attempts:
                logger.info("Retrying %s (next attempt %s)", getattr(callable_func, "__name__", str(callable_func)), attempt + 1)
    # Все попытки исчерпаны
    logger.error("All attempts failed for %s", getattr(callable_func, "__name__", str(callable_func)))
    print(last_exc)
    raise last_exc

def test(jj=None):
    fdsf = 'djfksl'
    if not pd.isna(jj):
        print('получилось')
        return
    try:
        return int(fdsf)
    except:
        attempt_recover(test())


logger.exception("CzMain: failed to initialize search/data/headers; using empty defaults")

#
# print(attempt_recover('ssdadsaasd'))

