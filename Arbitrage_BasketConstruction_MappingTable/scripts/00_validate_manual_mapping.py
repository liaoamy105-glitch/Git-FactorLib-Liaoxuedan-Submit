from pathlib import Path

import pandas as pd


INPUT_PATH = Path("data/manual/domestic_foreign_mapping.csv")
REPORT_PATH = Path("data/processed/manual_mapping_validation_report.csv")

REQUIRED_COLUMNS = [
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

NON_EMPTY_COLUMNS = [
    "domestic_product_code",
    "domestic_exchange_code",
    "foreign_product_code",
    "foreign_exchange_code",
    "mapping_order",
]


def add_issue(issues, issue_type, column, row_number, message):
    issues.append(
        {
            "status": "failed",
            "issue_type": issue_type,
            "column": column,
            "row_number": row_number,
            "message": message,
        }
    )


def is_empty(value):
    return pd.isna(value) or str(value).strip() == ""


def main():
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    issues = []

    if not INPUT_PATH.exists():
        add_issue(
            issues,
            "missing_file",
            "",
            "",
            f"Manual mapping file not found: {INPUT_PATH}",
        )
        pd.DataFrame(issues).to_csv(REPORT_PATH, index=False)
        raise SystemExit(1)

    mapping = pd.read_csv(INPUT_PATH, dtype=str)

    for column in REQUIRED_COLUMNS:
        if column not in mapping.columns:
            add_issue(
                issues,
                "missing_column",
                column,
                "",
                f"Required column is missing: {column}",
            )

    for column in NON_EMPTY_COLUMNS:
        if column not in mapping.columns:
            continue
        empty_rows = mapping[mapping[column].apply(is_empty)]
        for index in empty_rows.index:
            add_issue(
                issues,
                "empty_required_value",
                column,
                index + 2,
                f"Required value is empty: {column}",
            )

    duplicate_check_columns = [
        "domestic_product_code",
        "domestic_exchange_code",
        "mapping_order",
    ]
    if all(column in mapping.columns for column in duplicate_check_columns):
        duplicate_mask = mapping.duplicated(duplicate_check_columns, keep=False)
        duplicate_rows = mapping[duplicate_mask]
        for index, row in duplicate_rows.iterrows():
            add_issue(
                issues,
                "duplicate_mapping_order",
                "mapping_order",
                index + 2,
                "mapping_order is duplicated under the same domestic_product_code and domestic_exchange_code: "
                f"{row['domestic_product_code']} / {row['domestic_exchange_code']} / {row['mapping_order']}",
            )

    if issues:
        pd.DataFrame(issues).to_csv(REPORT_PATH, index=False)
        raise SystemExit(1)

    pd.DataFrame(
        [
            {
                "status": "passed",
                "issue_type": "",
                "column": "",
                "row_number": "",
                "message": "manual mapping validation passed",
            }
        ]
    ).to_csv(REPORT_PATH, index=False)
    print("manual mapping validation passed")


if __name__ == "__main__":
    main()
