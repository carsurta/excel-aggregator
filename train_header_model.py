
from __future__ import annotations
"""
train_header_model.py

./data/header_labels.csv  (UTF-8, BOM 없이 또는 with)
./data/*.xlsx             (원본 제출 양식들)

형식의 지도학습 데이터를 사용하여
"어떤 (start_row, end_row) 밴드가 헤더인가?"를 예측하는 모델을 학습한다.

CSV 포맷 (1행 헤더):

    file_name,sheet_name,header_start_1based,header_end_1based

예시:

    시험데이터_1.xlsx,제출용,3,4
    시험데이터_2.xlsx,제출본,2,3
    시험데이터_3.xlsx,제출용,3,5

학습 완료 후, 동일 폴더에 header_band_model.pkl 을 저장한다.
(이 파일은 나중에 header_multirow.py 또는 별도 모듈에서 로드하여 사용하면 된다.)
"""
import os
from pathlib import Path
import pickle

import numpy as np
import pandas as pd
from header_multirow import band_features, detect_data_start_strict
try:
    import lightgbm as lgb
    HAS_LGBM = True
except Exception:
    HAS_LGBM = False
    from sklearn.ensemble import HistGradientBoostingClassifier

from header_multirow import load_sheet_merge_aware, HEADER_SCAN_ROWS
from sklearn.metrics import accuracy_score, precision_recall_fscore_support


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
LABEL_CSV = DATA_DIR / "header_labels.csv"
OUT_PATH = BASE_DIR / "header_band_model.pkl"


def build_dataset() -> tuple[np.ndarray, np.ndarray, list[tuple[str, str, int, int]]]:
    if not LABEL_CSV.exists():
        raise FileNotFoundError(f"라벨 CSV를 찾을 수 없습니다: {LABEL_CSV}")

    labels = pd.read_csv(LABEL_CSV)
    required = {"file_name", "sheet_name", "header_start_1based", "header_end_1based"}
    if not required.issubset(labels.columns):
        raise ValueError(f"CSV에는 반드시 컬럼 {required} 가 포함되어야 합니다.")

    X: list[list[float]] = []
    y: list[int] = []
    meta: list[tuple[str, str, int, int]] = []

    for row in labels.itertuples(index=False):
        fname = getattr(row, "file_name")
        sname = getattr(row, "sheet_name")
        h_start = int(getattr(row, "header_start_1based")) - 1  # 0-based
        h_end = int(getattr(row, "header_end_1based")) - 1

        xlsx_path = DATA_DIR / fname
        if not xlsx_path.exists():
            print(f"[WARN] 파일 없음: {xlsx_path}")
            continue

        df = load_sheet_merge_aware(str(xlsx_path), sname)
        dstar = detect_data_start_strict(df)
        scan = min(HEADER_SCAN_ROWS, len(df))

        # 상단부에서 다양한 후보 밴드를 만들어 학습 데이터 구성
        search_limit = min(scan, max(12, dstar + 2))
        for r in range(0, search_limit):
            max_depth = min(5, search_limit - r)
            if max_depth <= 0:
                break
            for depth in range(1, max_depth + 1):
                end = r + depth - 1
                if abs(r - h_start) <= 1 and abs(end - h_end) <= 1 and not (r == h_start and end == h_end):
                    continue
                feats = band_features(df, r, end, dstar)
                label = 1 if (r == h_start and end == h_end) else 0
                X.append(feats)
                y.append(label)
                meta.append((fname, sname, r, end))

    if not X:
        return np.empty((0, 0), float), np.empty((0,), int), meta

    X = np.array(X, dtype=float)
    y = np.array(y, dtype=int)
    rng = np.random.default_rng(42)
    perm = rng.permutation(len(X))
    X = X[perm]
    y = y[perm]
    meta = [meta[i] for i in perm]

    return X, y, meta


def main():
    X, y, meta = build_dataset()
    if X.shape[0] == 0:
        print("학습 데이터가 비어 있습니다. data/header_labels.csv 와 data/*.xlsx 를 확인하세요.")
        return

    if len(set(y.tolist())) < 2:
        print("y 라벨이 한 종류 뿐입니다. header_start_1based / header_end_1based 값을 다시 확인하세요.")
        return

    if HAS_LGBM:
        print("Using LightGBM (supervised).")
        dtrain = lgb.Dataset(X, label=y)
        params = dict(
            objective="binary",
            boosting="gbdt",
            learning_rate=0.1,
            num_leaves=31,
            feature_fraction=0.9,
            bagging_fraction=0.8,
            bagging_freq=1,
            verbose=-1,
        )
        model = lgb.train(params, dtrain, num_boost_round=400)
        train_scores = np.ravel(model.predict(X))
        impl = ("lightgbm", model)
    else:
        print("LightGBM 미탑재: HistGradientBoostingClassifier 사용 (supervised).")
        model = HistGradientBoostingClassifier(
            max_depth=6,
            learning_rate=0.1,
            max_iter=400,
        )
        model.fit(X, y)
        proba = model.predict_proba(X)
        train_scores = proba[:, 1] if proba.ndim == 2 else np.ravel(proba)
        impl = ("sk_hgb", model)

    with open(OUT_PATH, "wb") as f:
        pickle.dump(impl, f)

    preds = (np.ravel(train_scores) >= 0.5).astype(int)
    acc = accuracy_score(y, preds)
    prec, recall, f1, _ = precision_recall_fscore_support(y, preds, average="binary", zero_division=0)
    pos = int(y.sum())
    neg = int((y == 0).sum())
    print(f"모델 저장 완료: {OUT_PATH}")
    print(f"총 샘플 수: {len(y)} (positive={pos}, negative={neg})")
    print(f"학습 지표 — acc: {acc:.3f}, precision: {prec:.3f}, recall: {recall:.3f}, f1: {f1:.3f}")


if __name__ == "__main__":
    main()
