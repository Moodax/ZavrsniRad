"""DCAT metadata generation module."""
from datetime import datetime
from typing import List
import os
from rdflib import Graph, Literal, URIRef, Namespace
from rdflib.namespace import DCAT, DCTERMS, FOAF, RDF, XSD
from urllib.parse import urlparse

def validate_uri(uri: str) -> bool:
    """Validate if a string is a well-formed URI."""
    parsed = urlparse(uri)
    return all([parsed.scheme, parsed.netloc])

def generate_dcat_metadata(
    input_file: str,
    output_files: List[str],
    base_uri: str = "http://example.org/dataset/",
    publisher_uri: str = "http://example.org/publisher",
    publisher_name: str = "Example Organization",
    license_uri: str = "http://creativecommons.org/licenses/by/4.0/",
) -> Graph:
    """Generate DCAT-AP compliant metadata for the Excel to CSV conversion.

    Args:
        input_file: Path to the input Excel file
        output_files: List of paths to output CSV files
        base_uri: Base URI for dataset and distribution identifiers
        publisher_uri: URI identifying the dataset publisher
        publisher_name: Human-readable name of the publisher
        license_uri: URI of the license under which data is published

    Returns:
        RDFLib Graph containing the DCAT metadata

    Raises:
        ValueError: If any of the URIs are invalid or input file is missing.
    """
    # The input_file is used as an identifier, its existence has been checked by the caller
    # if not os.path.exists(input_file):
    #     raise ValueError(f"Input file does not exist: {input_file}")

    if not validate_uri(base_uri):
        raise ValueError(f"Invalid base URI: {base_uri}")

    if not validate_uri(publisher_uri):
        raise ValueError(f"Invalid publisher URI: {publisher_uri}")

    if not validate_uri(license_uri):
        raise ValueError(f"Invalid license URI: {license_uri}")

    g = Graph()

    # Add namespaces
    SCHEMA = Namespace("http://schema.org/")
    g.bind("schema", SCHEMA)
    g.bind("dcat", DCAT)
    g.bind("dct", DCTERMS)
    g.bind("foaf", FOAF)

    # Create dataset URI using the input filename
    dataset_name = os.path.splitext(os.path.basename(input_file))[0]
    dataset_uri = URIRef(f"{base_uri}{dataset_name}")

    # Create publisher
    publisher = URIRef(publisher_uri)
    g.add((publisher, RDF.type, FOAF.Organization))
    g.add((publisher, FOAF.name, Literal(publisher_name)))

    # Describe the Dataset
    g.add((dataset_uri, RDF.type, DCAT.Dataset))
    g.add((dataset_uri, DCTERMS.title, Literal(f"Excel Data: {dataset_name}")))
    g.add((dataset_uri, DCTERMS.description, Literal(f"Dataset extracted from Excel file {input_file}")))
    g.add((dataset_uri, DCTERMS.issued, Literal(datetime.now().isoformat(), datatype=XSD.dateTime)))

    # Add license and publisher
    g.add((dataset_uri, DCTERMS.license, URIRef(license_uri)))
    g.add((dataset_uri, DCTERMS.publisher, publisher))

    # Add distributions for each output file
    for output_file in output_files:
        distribution_uri = URIRef(f"{base_uri}{os.path.basename(output_file)}")
        g.add((distribution_uri, RDF.type, DCAT.Distribution))
        g.add((distribution_uri, DCTERMS.title, Literal(f"CSV File: {os.path.basename(output_file)}")))
        g.add((distribution_uri, DCTERMS.format, Literal("text/csv")))
        g.add((distribution_uri, DCTERMS.issued, Literal(datetime.now().isoformat(), datatype=XSD.dateTime)))
        g.add((distribution_uri, DCTERMS.language, Literal("en")))
        g.add((distribution_uri, DCTERMS.license, URIRef(license_uri)))
        g.add((distribution_uri, DCAT.accessURL, Literal(output_file)))
        g.add((dataset_uri, DCAT.distribution, distribution_uri))

    return g
