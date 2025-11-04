import logging
from typing import Optional

from amex_tool.pipeline import (
    apply_corporate_cards,
    clean_statement,
    generate_claim_frames,
    load_corporate_mapping,
    load_statement,
    load_template_columns,
    save_claim_frames,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def main(
    statement_file: str,
    output_dir: str,
    corporate_file: Optional[str] = None,
    export_format: str = "excel",
    template_file: Optional[str] = None,
) -> None:
    statement_df = load_statement(statement_file)
    cleaned_df = clean_statement(statement_df)

    template_columns = None
    if template_file:
        template_columns = load_template_columns(template_file)

    claim_frames = generate_claim_frames(cleaned_df, template_columns)

    mapping = load_corporate_mapping(corporate_file) if corporate_file else None
    claim_frames = apply_corporate_cards(claim_frames, mapping)

    exported = save_claim_frames(claim_frames, output_dir, export_format)
    logging.info("✅ Processing complete. Generated %s files.", len(exported))


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Amex to Acumatica Claim File Processor"
    )
    parser.add_argument("--statement", required=True, help="Path to Amex statement file")
    parser.add_argument(
        "--output", required=True, help="Directory where claim files will be written"
    )
    parser.add_argument(
        "--corporate", help="Optional corporate card mapping file (CSV or Excel)"
    )
    parser.add_argument(
        "--template",
        help="Optional template file to define output column order (CSV or Excel)",
    )
    parser.add_argument(
        "--format",
        choices=["csv", "excel"],
        default="excel",
        help="Export format for claim files",
    )

    args = parser.parse_args()
    main(args.statement, args.output, args.corporate, args.format, args.template)
