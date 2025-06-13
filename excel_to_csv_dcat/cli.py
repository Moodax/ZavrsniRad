"""Command-line interface for Excel to CSV conversion with DCAT metadata."""

import sys
import os
import argparse
from typing import List, Tuple

from .core import process_excel_in_memory

def parse_args() -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        prog="excel-to-csv-dcat",
        description="Parse Excel tables and extract them to CSV files with DCAT-AP metadata"
    )

    parser.add_argument("filename", help="Input Excel file path")
    parser.add_argument(
        "-o", "--output-dir",
        default="output",
        help="Output directory for CSV files and metadata"
    )
    parser.add_argument(
        "-m", "--metadata-format",
        choices=["turtle", "json-ld"],
        default="turtle",
        help="Metadata output format"
    )
    parser.add_argument(
        "-b", "--base-uri",
        default="http://example.org/dataset/",
        help="Base URI for dataset and distribution identifiers"
    )
    parser.add_argument(
        "-p", "--publisher-uri",
        default="http://example.org/publisher",
        help="URI identifying the dataset publisher"
    )
    parser.add_argument(
        "-n", "--publisher-name",
        default="Example Organization",
        help="Human-readable name of the publisher"
    )
    parser.add_argument(
        "-l", "--license",
        default="http://creativecommons.org/licenses/by/4.0/",
        help="URI of the license under which data is published"
    )
    parser.add_argument(
        "--existing-dataset-uri",
        type=str,
        default=None,
        help="Optional URI of a pre-existing DCAT Dataset to link generated distributions to."
    )
    parser.add_argument(
        "--original-source-description",
        type=str,
        default=None,
        help="Optional text describing the original source of the dataset."
    )
    parser.add_argument(
        "--enable-ai",
        action="store_true",
        help="Enable AI-powered features for header generation and datatype validation"
    )
    parser.add_argument(
        "--llm-provider",
        choices=["openai", "gemini"],
        default="openai",
        help="LLM provider to use for AI features (default: openai)"
    )
    parser.add_argument(
        "--llm-api-key",
        type=str,
        help="API key for the selected LLM provider (can also be set via environment variables)"
    )
    parser.add_argument(
        "--skip-header-ai",
        action="store_true",
        help="Skip AI-powered header generation for headerless tables"
    )
    parser.add_argument(
        "--skip-datatype-ai",
        action="store_true",
        help="Skip AI-powered datatype validation"
    )

    return parser.parse_args()

def validate_file_path(file_path: str) -> None:
    """Validate that the input file exists and is accessible."""
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")

def validate_output_dir(output_dir: str) -> None:
    """Ensure the output directory exists or create it."""
    os.makedirs(output_dir, exist_ok=True)

def process_file(args: argparse.Namespace) -> Tuple[List[str], str]:
    """Process the Excel file according to command line arguments."""
    validate_file_path(args.filename)
    validate_output_dir(args.output_dir)

    # Read Excel file into memory
    with open(args.filename, 'rb') as f:
        excel_bytes = f.read()

    # Process the Excel file
    csv_files, metadata_buffer = process_excel_in_memory(
        excel_bytes,
        args.filename,
        args.output_dir,
        base_uri=args.base_uri,
        publisher_uri=args.publisher_uri,
        publisher_name=args.publisher_name,
        license_uri=args.license,
        existing_dataset_uri=args.existing_dataset_uri,
        original_source_description=args.original_source_description,
        enable_ai=args.enable_ai,
        llm_provider=args.llm_provider,
        llm_api_key=args.llm_api_key,
        skip_header_ai=args.skip_header_ai,
        skip_datatype_ai=args.skip_datatype_ai
    )

    # Save metadata
    metadata_file = os.path.join(args.output_dir, f"metadata.{args.metadata_format}")
    with open(metadata_file, 'wb') as f:
        f.write(metadata_buffer.getvalue())

    return [f["path"] for f in csv_files], metadata_file

def main() -> None:
    """Main entry point for the CLI."""
    try:
        args = parse_args()
        output_files, metadata_file = process_file(args)

        print(f"✓ Processed {len(output_files)} tables from '{args.filename}'")
        print(f"✓ CSV files saved to: {args.output_dir}")
        for output_file in output_files:
            print(f"  - {os.path.basename(output_file)}")
        print(f"✓ Metadata saved to: {metadata_file}")

    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
