"""Persistent on-disk cache for QUERY_BASE_DATA pulls, keyed by scope
(which client(s) + how many months back) so different selections don't
overwrite each other's cached data.
"""
import hashlib
import os
import time
from typing import Iterable, Optional

import pandas as pd

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")

DEFAULT_TTL_SECONDS = 60 * 60  # 1 hour


def make_key(company_ids: Optional[Iterable[int]], months_back: int) -> str:
    if company_ids:
        ids_part = ",".join(str(int(cid)) for cid in sorted(company_ids))
    else:
        ids_part = "ALL"
    raw = f"{ids_part}|{months_back}"
    return hashlib.md5(raw.encode()).hexdigest()[:16]


def _cache_file(key: str) -> str:
    return os.path.join(CACHE_DIR, f"query_base_data_{key}.pkl")


def save(df: pd.DataFrame, key: str) -> None:
    os.makedirs(CACHE_DIR, exist_ok=True)
    df.to_pickle(_cache_file(key))


def load(key: str) -> Optional[pd.DataFrame]:
    path = _cache_file(key)
    if not os.path.exists(path):
        return None
    try:
        return pd.read_pickle(path)
    except Exception:
        return None


def age_seconds(key: str) -> Optional[float]:
    path = _cache_file(key)
    if not os.path.exists(path):
        return None
    return time.time() - os.path.getmtime(path)


def is_fresh(key: str, ttl_seconds: int = DEFAULT_TTL_SECONDS) -> bool:
    age = age_seconds(key)
    return age is not None and age < ttl_seconds


def clear(key: str) -> None:
    path = _cache_file(key)
    if os.path.exists(path):
        os.remove(path)
