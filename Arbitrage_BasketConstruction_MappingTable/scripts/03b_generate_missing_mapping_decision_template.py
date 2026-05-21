from pathlib import Path

import pandas as pd


INPUT_PATH = Path("data/processed/missing_foreign_mapping_candidates.csv")
OUTPUT_PATH = Path("data/manual/domestic_foreign_mapping_decision.csv")

EXCLUDE_FROM_FINAL_CODES = {
    "ME",
    "RI",
    "RO",
    "WS",
    "WT",
    "L_F",
    "PP_F",
    "V_F",
    "SCTAS",
}

NO_FOREIGN_MAPPING_CODES = {
    "T",
    "TF",
    "TL",
    "TS",
}

OUTPUT_COLUMNS = [
    "domestic_product_code",
    "domestic_product_name",
    "domestic_exchange_code",
    "domestic_exchange_name",
    "decision",
    "reason",
    "foreign_product_code",
    "foreign_product_name",
    "foreign_exchange_code",
    "foreign_exchange_name",
    "foreign_country",
    "mapping_order",
    "mapping_strength",
    "mapping_note",
]


def normalize_code(value):
    return str(value).strip().upper()


def decision_for_product(product_code):
    normalized_code = normalize_code(product_code)

    if normalized_code in EXCLUDE_FROM_FINAL_CODES:
        return "EXCLUDE_FROM_FINAL", "historical_or_special_contract"

    if normalized_code in NO_FOREIGN_MAPPING_CODES:
        return "NO_FOREIGN_MAPPING", "no_clear_foreign_equivalent"

    return "NEED_CONFIRM", "empty"


def main():
    missing_candidates = pd.read_csv(INPUT_PATH, dtype=str).fillna("")

    rows = []
    for _, row in missing_candidates.iterrows():
        decision, reason = decision_for_product(row["domestic_product_code"])
        rows.append(
            {
                "domestic_product_code": row["domestic_product_code"],
                "domestic_product_name": row["domestic_product_name"],
                "domestic_exchange_code": row["domestic_exchange_code"],
                "domestic_exchange_name": row["domestic_exchange_name"],
                "decision": decision,
                "reason": reason,
                "foreign_product_code": "",
                "foreign_product_name": "",
                "foreign_exchange_code": "",
                "foreign_exchange_name": "",
                "foreign_country": "",
                "mapping_order": "",
                "mapping_strength": "",
                "mapping_note": "",
            }
        )

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows, columns=OUTPUT_COLUMNS).to_csv(
        OUTPUT_PATH, index=False, encoding="utf-8-sig"
    )

    print("missing mapping decision template generated")
    print(f"decision template saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
