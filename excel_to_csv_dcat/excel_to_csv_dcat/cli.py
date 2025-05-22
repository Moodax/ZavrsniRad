"""Command-line interface for Excel to CSV conversion with DCAT metadata."""

import sys
import os
import argparse
from typing import List, Tuple

from .core import process_excel_in_memory
from .metadata import generate_dcat_metadata

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
    output_files = process_excel_in_memory(excel_bytes, args.output_dir, args.filename)

    # Generate DCAT metadata
    g = generate_dcat_metadata(
        args.filename,
        output_files,
        base_uri=args.base_uri,
        publisher_uri=args.publisher_uri,
        publisher_name=args.publisher_name,
        license_uri=args.license
    )

    # Save metadata
    metadata_file = os.path.join(args.output_dir, f"metadata.{args.metadata_format}")
    g.serialize(destination=metadata_file, format=args.metadata_format)

    return output_files, metadata_file

def main() -> None:
    """Main entry point for the CLI."""
    args = parse_args()

    try:
        # Convert paths to absolute and normalize them
        input_file = os.path.abspath(os.path.normpath(args.filename))
        output_dir = os.path.abspath(os.path.normpath(args.output_dir))

        if not os.path.exists(input_file):
            print(f"Error: Input file not found: {input_file}", file=sys.stderr)
            sys.exit(1)

        # Read Excel file into memory
        with open(input_file, "rb") as f:
            excel_bytes = f.read()

        # Process Excel file and get CSV files and metadata
        csv_files, metadata_buffer = process_excel_in_memory(
            excel_bytes,
            os.path.basename(input_file), # Pass the original Excel filename
            output_dir,
            base_uri=args.base_uri,
            publisher_uri=args.publisher_uri,
            publisher_name=args.publisher_name,
            license_uri=args.license # Make sure 'license' is the correct attribute name from args
        )

        # Save metadata
        # Determine the correct extension based on args.metadata_format
        metadata_ext = args.metadata_format
        if args.metadata_format == "turtle":
            metadata_ext = "ttl"
        elif args.metadata_format == "json-ld":
            metadata_ext = "jsonld"
        
        metadata_file_path = os.path.join(output_dir, f"metadata.{metadata_ext}")
        
        with open(metadata_file_path, "wb") as f:
            f.write(metadata_buffer.getvalue())

        print(f"Successfully processed {len(csv_files)} tables:")
        for csv_file in csv_files:
            print(f"- {os.path.basename(csv_file)}")
        print(f"Metadata saved to: {os.path.basename(metadata_file_path)}")

    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        # Consider more specific error handling or logging
        # import traceback
        # traceback.print_exc() 
        sys.exit(1)

if __name__ == '__main__':
    main()