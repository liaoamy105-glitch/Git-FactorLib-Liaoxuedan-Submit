from pathlib import Path

import pandas as pd


DOMESTIC_MASTER_INPUT = Path("data/processed/domestic_master.csv")
MANUAL_MAPPING_INPUT = Path("data/manual/domestic_foreign_mapping.csv")
COVERAGE_REPORT_OUTPUT = Path("data/processed/foreign_mapping_coverage_report.csv")
MISSING_CANDIDATES_OUTPUT = Path(
    "data/processed/missing_foreign_mapping_candidates.csv"
)
JOIN_PREVIEW_OUTPUT = Path("data/processed/foreign_mapping_join_preview.csv")

COVERAGE_COLUMNS = [
    "domestic_product_code",
    "domestic_product_name",
    "domestic_exchange_code",
    "domestic_exchange_name",
    "has_foreign_mapping",
    "foreign_mapping_count",
    "foreign_mapping_summary",
    "mapping_strength_summary",
    "status",
    "message",
]

JOIN_PREVIEW_COLUMNS = [
    "domestic_product_code",
    "domestic_product_name",
    "domestic_exchange_code",
    "domestic_exchange_name",
    "foreign_product_code",
    "foreign_product_name",
    "foreign_exchange_code",
    "foreign_exchange_name",
    "foreign_country",
    "mapping_order",
    "mapping_strength",
    "mapping_note",
    "source_url",
]


def read_csv(path):
    if not path.exists():
        raise FileNotFoundError(f"Required input file not found: {path}")
    return pd.read_csv(path, dtype=str).fillna("")


def normalize_key(value):
    return str(value).strip().upper()


def clean_text(value):
    return str(value).strip()


def foreign_mapping_summary(mapping_rows):
    summaries = []
    for _, row in mapping_rows.iterrows():
        parts = [
            clean_text(row["foreign_product_code"]),
            clean_text(row["foreign_exchange_code"]),
            clean_text(row["foreign_country"]),
        ]
        summaries.append("_".join(parts))
    return "; ".join(summaries)


def mapping_strength_summary(mapping_rows):
    strengths = [
        clean_text(value)
        for value in mapping_rows["mapping_strength"].tolist()
        if clean_text(value)
    ]
    return "; ".join(strengths)


def build_coverage_report(domestic_futures, manual_mapping):
    mapping_groups = {
        key: group
        for key, group in manual_mapping.groupby(
            ["_domestic_product_code_key", "_domestic_exchange_code_key"],
            dropna=False,
        )
    }

    rows = []
    for _, domestic in domestic_futures.iterrows():
        key = (
            domestic["_product_code_key"],
            domestic["_exchange_code_key"],
        )
        matched = mapping_groups.get(key, pd.DataFrame())
        has_mapping = not matched.empty
        mapping_count = len(matched)

        rows.append(
            {
                "domestic_product_code": domestic["product_code"],
                "domestic_product_name": domestic["product_name"],
                "domestic_exchange_code": domestic["exchange_code"],
                "domestic_exchange_name": domestic["exchange_name"],
                "has_foreign_mapping": has_mapping,
                "foreign_mapping_count": mapping_count,
                "foreign_mapping_summary": foreign_mapping_summary(matched)
                if has_mapping
                else "",
                "mapping_strength_summary": mapping_strength_summary(matched)
                if has_mapping
                else "",
                "status": "mapped" if has_mapping else "missing_need_confirm",
                "message": "foreign mapping found from manual file"
                if has_mapping
                else "no foreign mapping in manual file, need confirm whether to add mapping or keep blank",
            }
        )

    return pd.DataFrame(rows, columns=COVERAGE_COLUMNS)


def build_join_preview(domestic_futures, manual_mapping):
    joined = domestic_futures.merge(
        manual_mapping,
        left_on=["_product_code_key", "_exchange_code_key"],
        right_on=["_domestic_product_code_key", "_domestic_exchange_code_key"],
        how="inner",
        suffixes=("_master", "_manual"),
    )

    if joined.empty:
        return pd.DataFrame(columns=JOIN_PREVIEW_COLUMNS)

    preview = pd.DataFrame(
        {
            "domestic_product_code": joined["product_code"],
            "domestic_product_name": joined["product_name"],
            "domestic_exchange_code": joined["exchange_code"],
            "domestic_exchange_name": joined["exchange_name"],
            "foreign_product_code": joined["foreign_product_code"],
            "foreign_product_name": joined["foreign_product_name"],
            "foreign_exchange_code": joined["foreign_exchange_code"],
            "foreign_exchange_name": joined["foreign_exchange_name"],
            "foreign_country": joined["foreign_country"],
            "mapping_order": joined["mapping_order"],
            "mapping_strength": joined["mapping_strength"],
            "mapping_note": joined["mapping_note"],
            "source_url": joined["source_url"],
        }
    )

    return preview.reindex(columns=JOIN_PREVIEW_COLUMNS)


def main():
    domestic_master = read_csv(DOMESTIC_MASTER_INPUT)
    manual_mapping = read_csv(MANUAL_MAPPING_INPUT)

    domestic_futures = domestic_master[
        domestic_master["instrument_type"].str.strip().str.upper() == "FUTURE"
    ].copy()
    domestic_futures["_product_code_key"] = domestic_futures["product_code"].apply(
        normalize_key
    )
    domestic_futures["_exchange_code_key"] = domestic_futures["exchange_code"].apply(
        normalize_key
    )

    manual_mapping["_domestic_product_code_key"] = manual_mapping[
        "domestic_product_code"
    ].apply(normalize_key)
    manual_mapping["_domestic_exchange_code_key"] = manual_mapping[
        "domestic_exchange_code"
    ].apply(normalize_key)

    coverage_report = build_coverage_report(domestic_futures, manual_mapping)
    missing_candidates = coverage_report[
        coverage_report["has_foreign_mapping"] == False
    ].copy()
    join_preview = build_join_preview(domestic_futures, manual_mapping)

    COVERAGE_REPORT_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    coverage_report.to_csv(COVERAGE_REPORT_OUTPUT, index=False, encoding="utf-8-sig")
    missing_candidates.to_csv(
        MISSING_CANDIDATES_OUTPUT, index=False, encoding="utf-8-sig"
    )
    join_preview.to_csv(JOIN_PREVIEW_OUTPUT, index=False, encoding="utf-8-sig")

    domestic_checked = len(coverage_report)
    mapped_count = int(coverage_report["has_foreign_mapping"].sum())
    missing_count = domestic_checked - mapped_count

    print("foreign mapping coverage check completed")
    print(f"domestic futures checked: {domestic_checked}")
    print(f"mapped futures: {mapped_count}")
    print(f"missing futures: {missing_count}")
    print(f"coverage report saved to {COVERAGE_REPORT_OUTPUT}")
    print(f"missing candidates saved to {MISSING_CANDIDATES_OUTPUT}")
    print(f"join preview saved to {JOIN_PREVIEW_OUTPUT}")


if __name__ == "__main__":
    main()
