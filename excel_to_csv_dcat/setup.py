from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="excel_to_csv_dcat",
    version="0.1.0",
    author="Your Name", # Add your name
    author_email="your.email@example.com", # Add your email
    description="Excel to CSV converter with table detection and DCAT metadata generation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Moodax/ZavrsniRad", # Add your project URL
    packages=find_packages(),
    install_requires=[
        "pandas>=2.0",
        "openpyxl>=3.1",
        "rdflib>=7.0"
    ],
    entry_points={
        'console_scripts': [
            'excel_to_csv_dcat=excel_to_csv_dcat.cli:main',
            'excel_to_csv_dcat_gui=excel_to_csv_dcat.gui:main',
        ],
    },
)
