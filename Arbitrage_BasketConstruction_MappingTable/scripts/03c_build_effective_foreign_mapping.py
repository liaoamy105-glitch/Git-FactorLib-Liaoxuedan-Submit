from datetime import date
from pathlib import Path

import pandas as pd


BASE_MAPPING_INPUT = Path("data/manual/domestic_foreign_mapping.csv")
DECISION_INPUT = Path("data/manual/domestic_foreign_mapping_decision.csv")
DOMESTIC_MASTER_INPUT = Path("data/processed/domestic_master.csv")

EFFECTIVE_MAPPING_OUTPUT = Path("data/processed/effective_foreign_mapping.csv")
FINAL_SCOPE_OUTPUT = Path("data/processed/final_product_scope.csv")
SUMMARY_OUTPUT = Path("data/processed/effective_foreign_mapping_summary.csv")

MAPPING_COLUMNS = [
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
    "updated_at",
]

CODE_COLUMNS = [
    "domestic_product_code",
    "domestic_exchange_code",
    "foreign_product_code",
    "foreign_exchange_code",
]

DEDUP_COLUMNS = [
    "domestic_product_code",
    "domestic_exchange_code",
    "foreign_product_code",
    "foreign_exchange_code",
]

FINAL_SCOPE_COLUMNS = [
    "domestic_product_code",
    "domestic_product_name",
    "domestic_exchange_code",
    "domestic_exchange_name",
    "include_in_final",
    "foreign_mapping_policy",
    "scope_note",
]


def read_csv(path):
    if not path.exists():
        raise FileNotFoundError(f"Required input file not found: {path}")
    return pd.read_csv(path, dtype=str).fillna("")


def normalize_code(value):
    return str(value).strip().upper()


def normalize_mapping_codes(mapping):
    normalized = mapping.copy()
    for column in CODE_COLUMNS:
        normalized[column] = normalized[column].apply(normalize_code)
    return normalized


def decision_key(row):
    return (
        normalize_code(row["domestic_product_code"]),
        normalize_code(row["domestic_exchange_code"]),
    )


def build_add_mapping(decision):
    add_mapping = decision[
        decision["decision"].str.strip().str.upper() == "ADD_MAPPING"
    ].copy()

    if add_mapping.empty:
        return pd.DataFrame(columns=MAPPING_COLUMNS)

    today = date.today().strftime("%Y-%m-%d")
    add_mapping["source_url"] = ""
    add_mapping["updated_at"] = today

    return add_mapping.reindex(columns=MAPPING_COLUMNS).fillna("")


def build_effective_mapping(base_mapping, add_mapping):
    combined = pd.concat(
        [
            base_mapping.reindex(columns=MAPPING_COLUMNS).fillna(""),
            add_mapping.reindex(columns=MAPPING_COLUMNS).fillna(""),
        ],
        ignore_index=True,
    )
    combined = normalize_mapping_codes(combined)
    combined = combined.drop_duplicates(subset=DEDUP_COLUMNS, keep="first")
    return combined.reindex(columns=MAPPING_COLUMNS)


def build_final_scope(domestic_master, decision, effective_mapping):
    domestic_futures = domestic_master[
        domestic_master["instrument_type"].str.strip().str.upper() == "FUTURE"
    ].copy()
    domestic_futures["domestic_product_code"] = domestic_futures["product_code"].apply(
        normalize_code
    )
    domestic_futures["domestic_exchange_code"] = domestic_futures[
        "exchange_code"
    ].apply(normalize_code)

    decision_lookup = {}
    for _, row in decision.iterrows():
        key = decision_key(row)
        decision_lookup[key] = {
            "decision": str(row["decision"]).strip().upper(),
            "reason": str(row["reason"]).strip(),
        }

    mapped_keys = set(
        zip(
            effective_mapping["domestic_product_code"].apply(normalize_code),
            effective_mapping["domestic_exchange_code"].apply(normalize_code),
        )
    )

    rows = []
    for _, future in domestic_futures.iterrows():
        key = (
            future["domestic_product_code"],
            future["domestic_exchange_code"],
        )
        decision_record = decision_lookup.get(key)

        if decision_record and decision_record["decision"] == "EXCLUDE_FROM_FINAL":
            include_in_final = False
            policy = "EXCLUDE_FROM_FINAL"
            scope_note = decision_record["reason"]
        elif decision_record and decision_record["decision"] == "NO_FOREIGN_MAPPING":
            include_in_final = True
            policy = "NO_FOREIGN_MAPPING"
            scope_note = decision_record["reason"]
        elif decision_record and decision_record["decision"] == "ADD_MAPPING":
            include_in_final = True
            policy = "ADD_MAPPING"
            scope_note = decision_record["reason"]
        elif key in mapped_keys:
            include_in_final = True
            policy = "HAS_FOREIGN_MAPPING"
            scope_note = "foreign mapping found"
        else:
            include_in_final = True
            policy = "UNRESOLVED"
            scope_note = "not found in mapping or decision file"

        rows.append(
            {
                "domestic_product_code": future["domestic_product_code"],
                "domestic_product_name": future["product_name"],
                "domestic_exchange_code": future["domestic_exchange_code"],
                "domestic_exchange_name": future["exchange_name"],
                "include_in_final": include_in_final,
                "foreign_mapping_policy": policy,
                "scope_note": scope_note,
            }
        )

    return pd.DataFrame(rows, columns=FINAL_SCOPE_COLUMNS)


def build_summary(base_mapping, add_mapping, effective_mapping, final_scope):
    include_true = int(final_scope["include_in_final"].sum())
    exclude_false = len(final_scope) - include_true
    unresolved_count = int(
        (final_scope["foreign_mapping_policy"] == "UNRESOLVED").sum()
    )

    return pd.DataFrame(
        [
            {"metric": "base_mapping_rows", "value": len(base_mapping)},
            {"metric": "add_mapping_rows", "value": len(add_mapping)},
            {"metric": "effective_mapping_rows", "value": len(effective_mapping)},
            {"metric": "final_scope_total_futures", "value": len(final_scope)},
            {"metric": "final_scope_include_true", "value": include_true},
            {"metric": "final_scope_exclude_false", "value": exclude_false},
            {"metric": "final_scope_unresolved_count", "value": unresolved_count},
        ],
        columns=["metric", "value"],
    )


def main():
    base_mapping = read_csv(BASE_MAPPING_INPUT)
    decision = read_csv(DECISION_INPUT)
    domestic_master = read_csv(DOMESTIC_MASTER_INPUT)

    add_mapping = build_add_mapping(decision)
    effective_mapping = build_effective_mapping(base_mapping, add_mapping)
    final_scope = build_final_scope(domestic_master, decision, effective_mapping)
    summary = build_summary(base_mapping, add_mapping, effective_mapping, final_scope)

    EFFECTIVE_MAPPING_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    effective_mapping.to_csv(
        EFFECTIVE_MAPPING_OUTPUT, index=False, encoding="utf-8-sig"
    )
    final_scope.to_csv(FINAL_SCOPE_OUTPUT, index=False, encoding="utf-8-sig")
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    unresolved_count = int(
        (final_scope["foreign_mapping_policy"] == "UNRESOLVED").sum()
    )

    print("effective foreign mapping built successfully")
    print(f"effective mapping rows: {len(effective_mapping)}")
    print(f"final product scope rows: {len(final_scope)}")
    print(f"unresolved futures: {unresolved_count}")


if __name__ == "__main__":
    main()
