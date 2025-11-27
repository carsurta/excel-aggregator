
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
from sklearn.model_selection import train_test_split


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
LABEL_CSV = DATA_DIR / "header_labels.csv"
OUT_PATH = BASE_DIR / "header_band_model.pkl"

NEGATIVE_FACTOR = 8           # keep at most this multiple of negatives vs positives
POSITIVE_AUG_MULT = 3         # total copies of positive samples with noise
POSITIVE_AUG_NOISE = 0.02

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

    pos_idx = np.where(y == 1)[0]
    neg_idx = np.where(y == 0)[0]
    if len(pos_idx) == 0:
        return np.empty((0, 0), float), np.empty((0,), int), meta

    max_neg = min(len(neg_idx), NEGATIVE_FACTOR * len(pos_idx))
    if len(neg_idx) > max_neg:
        keep_neg = rng.choice(neg_idx, size=max_neg, replace=False)
    else:
        keep_neg = neg_idx
    keep_idx = np.concatenate([pos_idx, keep_neg])
    X = X[keep_idx]
    y = y[keep_idx]
    meta = [meta[i] for i in keep_idx]

    if POSITIVE_AUG_MULT > 1 and len(pos_idx) > 0:
        pos_mask = y == 1
        base_pos = X[pos_mask]
        base_meta = [meta[i] for i, lbl in enumerate(y) if lbl == 1]
        aug_X = [X]
        aug_y = [y]
        aug_meta = meta[:]
        for _ in range(POSITIVE_AUG_MULT - 1):
            noise = rng.normal(0, POSITIVE_AUG_NOISE, size=base_pos.shape)
            aug_X.append(base_pos + noise)
            aug_y.append(np.ones(len(base_pos), dtype=int))
            aug_meta.extend(base_meta)
        X = np.vstack(aug_X)
        y = np.concatenate(aug_y)
        meta = aug_meta

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

    pos_count = int((y == 1).sum())
    neg_count = int((y == 0).sum())
    scale_pos_weight = max(1.0, neg_count / max(1, pos_count))

    has_valid = len(y) >= 20 and neg_count > 0 and pos_count > 1
    if has_valid:
        X_train, X_valid, y_train, y_valid = train_test_split(
            X, y, test_size=0.2, random_state=42, stratify=y
        )
    else:
        X_train, y_train = X, y
        X_valid = y_valid = None

    if HAS_LGBM:
        print("Using LightGBM (supervised).")
        dtrain = lgb.Dataset(X_train, label=y_train)
        valid_sets = [dtrain]
        valid_names = ["train"]
        if X_valid is not None:
            dvalid = lgb.Dataset(X_valid, label=y_valid)
            valid_sets.append(dvalid)
            valid_names.append("valid")
        params = dict(
            objective="binary",
            boosting="gbdt",
            learning_rate=0.05,
            num_leaves=16,
            max_depth=6,
            min_data_in_leaf=5,
            min_split_gain=0.01,
            feature_fraction=0.85,
            bagging_fraction=0.7,
            bagging_freq=1,
            lambda_l1=0.1,
            lambda_l2=1.0,
            scale_pos_weight=scale_pos_weight,
            verbose=-1,
        )
        callbacks = []
        if X_valid is not None:
            callbacks.append(lgb.early_stopping(80, verbose=False))
        callbacks.append(lgb.log_evaluation(period=0))
        model = lgb.train(
            params,
            dtrain,
            num_boost_round=1200,
            valid_sets=valid_sets,
            valid_names=valid_names,
            callbacks=callbacks,
        )
        train_scores = np.ravel(model.predict(X))
        impl = ("lightgbm", model)
    else:
        print("LightGBM 미탑재: HistGradientBoostingClassifier 사용 (supervised).")
        model = HistGradientBoostingClassifier(
            max_depth=6,
            learning_rate=0.1,
            max_iter=400,
            class_weight="balanced",
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
    print(f"모델 저장 완료: {OUT_PATH}")
    print(f"총 샘플 수: {len(y)} (positive={pos_count}, negative={neg_count}, scale_pos_weight={scale_pos_weight:.2f})")
    print(f"학습 지표 — acc: {acc:.3f}, precision: {prec:.3f}, recall: {recall:.3f}, f1: {f1:.3f}")


if __name__ == "__main__":
    main()
