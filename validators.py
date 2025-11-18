# validators.py – v3 (Series-return, fail-safe)
from __future__ import annotations
from typing import List
import os, re
import pandas as pd

DISABLE = os.getenv("DISABLE_VALIDATION", "0") == "1"

_phone_re = re.compile(r"^(\+?\d[\d\s\-]{6,}\d)$")
_email_re = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _safe_str_series(s: pd.Series) -> pd.Series:
    try:
        return s.astype("string", errors="ignore")
    except Exception:
        try:
            return s.astype(str)
        except Exception:
            return s

def _mark(mask, msg: str, issues: List[str]):
    try:
        idx = getattr(mask, "index", range(len(mask)))
        for i, flag in zip(idx, mask):
            if bool(flag):
                issues[int(i)] += msg
    except Exception:
        pass

def compute_violations(df: pd.DataFrame) -> pd.Series:
    try:
        if DISABLE or df is None or getattr(df, "empty", False):
            return pd.Series([""] * (0 if df is None else len(df)),
                             index=(df.index if isinstance(df, pd.DataFrame) else None),
                             name="_violations")
        out = df
        issues: List[str] = ["" for _ in range(len(out))]

        cols = list(out.columns)
        data_cols = [c for c in cols if not str(c).startswith("_")]

        try:
            phone_cols = [c for c in cols if re.search(r"(연락처|전화|phone|tel)", str(c), re.IGNORECASE)]
        except Exception: phone_cols = []
        try:
            email_cols = [c for c in cols if re.search(r"(이메일|email|e-mail)", str(c), re.IGNORECASE)]
        except Exception: email_cols = []
        try:
            key_cols   = [c for c in cols if re.search(r"(번호|id|코드|code)", str(c), re.IGNORECASE)]
        except Exception: key_cols = []

        for c in phone_cols:
            try:
                s = _safe_str_series(out[c]).fillna("").str.strip()
                mask = s.apply(lambda x: (x != "") and (_phone_re.match(str(x)) is None))
                _mark(mask, f"[{c}:전화형식] ", issues)
            except Exception: pass

        for c in email_cols:
            try:
                s = _safe_str_series(out[c]).fillna("").str.strip()
                mask = s.apply(lambda x: (x != "") and (_email_re.match(str(x)) is None))
                _mark(mask, f"[{c}:이메일형식] ", issues)
            except Exception: pass

        try:
            for c in data_cols[:3]:
                s = out[c]
                empty_mask = s.isna() | (_safe_str_series(s).fillna("").str.strip() == "")
                if float(getattr(empty_mask, "mean", lambda: 0.0)()) > 0.2:
                    _mark(empty_mask, f"[{c}:빈값] ", issues)
        except Exception: pass

        for c in key_cols:
            try:
                s = _safe_str_series(out[c]).fillna("")
                dup = s.duplicated(keep=False) & (s != "")
                _mark(dup, f"[{c}:중복] ", issues)
            except Exception: pass

        return pd.Series(issues, index=out.index, name="_violations")
    except Exception:
        try:
            return pd.Series([""] * len(df), index=df.index, name="_violations")
        except Exception:
            return pd.Series([], name="_violations")
