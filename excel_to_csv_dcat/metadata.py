"""DCAT metadata generation module."""
from datetime import datetime
from typing import List, Optional, Dict
import os
import pandas as pd
import re
from rdflib import Graph, Literal, URIRef, Namespace
from rdflib.namespace import DCAT, DCTERMS, FOAF, RDF, XSD
from urllib.parse import urlparse, quote

# Add CSVW namespace
CSVW = Namespace("http://www.w3.org/ns/csvw#")

def validate_uri(uri: str) -> bool:
    """Validate if a string is a well-formed URI."""
    parsed = urlparse(uri)
    return all([parsed.scheme, parsed.netloc])

def map_pandas_dtype_to_xsd(dtype_obj) -> URIRef:
    """Maps a pandas dtype object to an appropriate XSD URIRef."""
    dtype_str = str(dtype_obj).lower()
    if "int" in dtype_str: 
        return XSD.integer
    if "float" in dtype_str or "decimal" in dtype_str: 
        return XSD.decimal
    if "bool" in dtype_str: 
        return XSD.boolean
    if "datetime" in dtype_str: 
        return XSD.dateTime
    if "date" in dtype_str: 
        return XSD.date
    if "time" in dtype_str: 
        return XSD.time
    return XSD.string  # Default for 'object', 'string', or unmapped types

def generate_dcat_metadata(
    input_file: str,
    saved_csv_details: List[dict],
    existing_dataset_uri: Optional[str] = None,
    original_source_description: Optional[str] = None,
    base_uri: str = "http://example.org/dataset/",
    publisher_uri: str = "http://example.org/publisher",
    publisher_name: str = "Example Organization",
    license_uri: str = "http://creativecommons.org/licenses/by/4.0/",
) -> Graph:
    """Generate DCAT-AP compliant metadata for the Excel to CSV conversion.

    Args:
        input_file: Path to the input Excel file
        saved_csv_details: List of dictionaries with CSV file details including titles
        existing_dataset_uri: Optional URI of a pre-existing dataset to link distributions to.
                              If provided, minimal dataset information will be generated.
        original_source_description: Optional text describing the original source of the dataset.
        base_uri: Base URI for dataset and distribution identifiers
        publisher_uri: URI identifying the dataset publisher
        publisher_name: Human-readable name of the publisher
        license_uri: URI of the license under which data is published

    Returns:
        RDFLib Graph containing the DCAT metadata

    Raises:
        ValueError: If any of the URIs are invalid or input file is missing.
    """
    # Validate URIs
    if not validate_uri(base_uri):
        raise ValueError(f"Invalid base URI: {base_uri}")

    if publisher_uri and not validate_uri(publisher_uri):
        raise ValueError(f"Invalid publisher URI: {publisher_uri}")

    if license_uri and not validate_uri(license_uri):
        raise ValueError(f"Invalid license URI: {license_uri}")

    g = Graph()

    # Add namespaces
    SCHEMA = Namespace("http://schema.org/")
    g.bind("schema", SCHEMA)
    g.bind("dcat", DCAT)
    g.bind("dct", DCTERMS)
    g.bind("foaf", FOAF)
    g.bind("csvw", CSVW)

    dataset_uri_to_use: URIRef  # Declare type for clarity

    if existing_dataset_uri and validate_uri(existing_dataset_uri):
        dataset_uri_to_use = URIRef(existing_dataset_uri)
        g.add((dataset_uri_to_use, RDF.type, DCAT.Dataset))  # Assert type for existing dataset
        
        # Create publisher FOAF.Organization if publisher_uri is provided, as it describes the agent doing the work.
        if publisher_uri and validate_uri(publisher_uri):
            publisher_agent_uri = URIRef(publisher_uri)
            g.add((publisher_agent_uri, RDF.type, FOAF.Organization))
            g.add((publisher_agent_uri, FOAF.name, Literal(publisher_name)))

    else:
        if existing_dataset_uri:  # It was provided but invalid
            print(f"Warning: Provided existing_dataset_uri '{existing_dataset_uri}' is invalid. Creating a new dataset description.")

        # Original behavior: create and fully describe a new dataset
        dataset_name = os.path.splitext(os.path.basename(input_file))[0]
        dataset_uri_to_use = URIRef(f"{base_uri}{dataset_name}")

        g.add((dataset_uri_to_use, RDF.type, DCAT.Dataset))
        g.add((dataset_uri_to_use, DCTERMS.title, Literal(f"Dataset derived from: {dataset_name}")))
        
        # Use original_source_description if available for the new dataset's description
        new_dataset_description_parts = [f"Dataset extracted from Excel file {input_file}."]
        if original_source_description:
            new_dataset_description_parts.append(f"Original source context: {original_source_description}.")
        g.add((dataset_uri_to_use, DCTERMS.description, Literal(" ".join(new_dataset_description_parts))))
        
        g.add((dataset_uri_to_use, DCTERMS.issued, Literal(datetime.now().isoformat(), datatype=XSD.dateTime)))

        # Add license and publisher for the NEW dataset
        if license_uri and validate_uri(license_uri):
            g.add((dataset_uri_to_use, DCTERMS.license, URIRef(license_uri)))
        
        if publisher_uri and validate_uri(publisher_uri):
            publisher_org_uri = URIRef(publisher_uri)
            g.add((publisher_org_uri, RDF.type, FOAF.Organization))
            g.add((publisher_org_uri, FOAF.name, Literal(publisher_name)))
            g.add((dataset_uri_to_use, DCTERMS.publisher, publisher_org_uri))    # Add distributions for each CSV file
    for csv_detail in saved_csv_details:
        output_file = os.path.basename(csv_detail["path"])
        distribution_uri = URIRef(f"{base_uri}{quote(output_file)}")
        g.add((distribution_uri, RDF.type, DCAT.Distribution))
        
        # Create a unique URI for the table schema associated with this distribution
        schema_uri = URIRef(str(distribution_uri) + "#schema")
        
        # Link the distribution to this table schema
        g.add((distribution_uri, CSVW.tableSchema, schema_uri))
        
        # Assert the type of the schema resource
        g.add((schema_uri, RDF.type, CSVW.Schema))
        
        # Use extracted title if available, otherwise default title
        title = csv_detail.get("table_title")
        if not title:
            title = f"CSV File: {output_file}"
        
        g.add((distribution_uri, DCTERMS.title, Literal(title)))
        
        # Enhanced description for distribution using original_source_description
        dist_title_for_desc = title
        input_filename_basename = os.path.basename(input_file)
        
        distribution_description_parts = [f"Table '{dist_title_for_desc}' (CSV format) extracted from the Excel file '{input_filename_basename}'."]
        
        if original_source_description:
            distribution_description_parts.append(f"This data originates from a source described as: '{original_source_description}'.")
        else:
            distribution_description_parts.append("The original source of this data was not specified.")
            
        final_distribution_description = " ".join(distribution_description_parts)
        g.add((distribution_uri, DCTERMS.description, Literal(final_distribution_description)))
        
        g.add((distribution_uri, DCTERMS.format, Literal("text/csv")))
        g.add((distribution_uri, DCTERMS.issued, Literal(datetime.now().isoformat(), datatype=XSD.dateTime)))
        g.add((distribution_uri, DCTERMS.language, Literal("en")))
        if license_uri:
            g.add((distribution_uri, DCTERMS.license, URIRef(license_uri)))
        g.add((distribution_uri, DCAT.downloadURL, Literal(output_file)))  # Changed from accessURL to downloadURL
        g.add((dataset_uri_to_use, DCAT.distribution, distribution_uri))  # Use dataset_uri_to_use
        
        # CSVW Column Processing with Type Inference
        column_headers = csv_detail.get("headers", [])
        csv_file_path = csv_detail["path"]

        # Dictionary to store inferred XSD types for each column header
        inferred_column_types: Dict[str, URIRef] = {}
        
        # Check if AI-validated column types are available
        validated_column_types = csv_detail.get("validated_column_types", {})

        if os.path.exists(csv_file_path) and column_headers:
            try:
                # Read a sample of the CSV for type inference
                df_sample = pd.read_csv(csv_file_path, nrows=100)
                
                for header_name_from_df in df_sample.columns:
                    original_header_name = str(header_name_from_df)
                    
                    # Use AI-validated type if available, otherwise infer from pandas
                    if original_header_name in validated_column_types:
                        # Convert XSD string to URIRef
                        xsd_type_str = validated_column_types[original_header_name]
                        if xsd_type_str.startswith("xsd:"):
                            xsd_type_local = xsd_type_str[4:]  # Remove "xsd:" prefix
                            inferred_column_types[original_header_name] = getattr(XSD, xsd_type_local, XSD.string)
                        else:
                            inferred_column_types[original_header_name] = XSD.string
                    else:
                        # Fallback to pandas dtype inference
                        inferred_column_types[original_header_name] = map_pandas_dtype_to_xsd(df_sample[original_header_name].dtype)

            except pd.errors.EmptyDataError:
                print(f"Warning: CSV file '{csv_file_path}' is empty. Cannot infer column types.")
                for header_name in column_headers:
                    inferred_column_types[str(header_name)] = XSD.string
            except FileNotFoundError:
                print(f"Warning: CSV file '{csv_file_path}' not found. Cannot infer column types.")
                for header_name in column_headers:
                    inferred_column_types[str(header_name)] = XSD.string
            except Exception as e:
                print(f"Warning: Error reading or inferring types for '{csv_file_path}': {e}. Defaulting types to string.")
                for header_name in column_headers:
                    inferred_column_types[str(header_name)] = XSD.string
        else:
            for header_name in column_headers:
                 inferred_column_types[str(header_name)] = XSD.string

        # Generate CSVW column triples
        for col_idx, original_header_name_obj in enumerate(column_headers):
            original_header_name = str(original_header_name_obj)

            # Create a URL-safe name for the column URI fragment
            safe_col_name_for_uri_part = re.sub(r'[^a-zA-Z0-9_]', '_', original_header_name)
            if not re.search(r'[a-zA-Z0-9]', safe_col_name_for_uri_part):
                safe_col_name_for_uri_part = f"column_{col_idx + 1}"
            
            column_uri = URIRef(str(schema_uri) + f"/{safe_col_name_for_uri_part}")

            g.add((schema_uri, CSVW.column, column_uri))
            g.add((column_uri, RDF.type, CSVW.Column))
            
            # CSVW.name is for the actual name in the CSV file
            g.add((column_uri, CSVW.name, Literal(original_header_name)))
            
            # DCTERMS.title can be a more human-readable version
            g.add((column_uri, DCTERMS.title, Literal(original_header_name)))
            
            # Get the inferred XSD datatype for this column
            xsd_datatype = inferred_column_types.get(original_header_name, XSD.string)
            g.add((column_uri, CSVW.datatype, xsd_datatype))

    return g
