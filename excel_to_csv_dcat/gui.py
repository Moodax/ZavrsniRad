"""Graphical user interface for Excel to CSV converter with DCAT metadata."""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional

from .core import process_excel_in_memory
from .metadata import generate_dcat_metadata

def validate_input_file(file_path: Optional[str]) -> None:
    """Validate that the input file is selected and exists."""
    if not file_path or not os.path.isfile(file_path):
        raise FileNotFoundError("Please select a valid input file.")

def validate_output_dir(output_dir: Optional[str]) -> None:
    """Validate that the output directory is selected and accessible."""
    if not output_dir:
        raise ValueError("Please select an output directory.")
    os.makedirs(output_dir, exist_ok=True)

class ExcelConverterGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel to CSV Converter with DCAT")

        self.root.geometry("600x400")

        # Variables
        self.input_file: Optional[str] = None
        self.output_dir: Optional[str] = None
        self.base_uri = tk.StringVar(value="http://example.org/dataset/") # Added base_uri
        self.publisher_name = tk.StringVar(value="Example Organization")
        self.publisher_uri = tk.StringVar(value="http://example.org/publisher")
        self.license_uri = tk.StringVar(value="http://creativecommons.org/licenses/by/4.0/")

        # AI-related variables
        self.enable_ai = tk.BooleanVar(value=False)
        self.llm_provider = tk.StringVar(value="openai")
        self.llm_api_key = tk.StringVar(value="")
        self.skip_header_ai = tk.BooleanVar(value=False)
        self.skip_datatype_ai = tk.BooleanVar(value=False)

        self._create_widgets()

    def _create_widgets(self):
        """Create and layout GUI widgets."""
        # Input file selection
        input_frame = ttk.LabelFrame(self.root, text="Input", padding="5")
        input_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        self.input_label = ttk.Label(input_frame, text="No file selected")
        self.input_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        ttk.Button(input_frame, text="Select Excel File", command=self._select_input).grid(
            row=0, column=1, padx=5, pady=5
        )

        # Output directory selection
        output_frame = ttk.LabelFrame(self.root, text="Output", padding="5")
        output_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        self.output_label = ttk.Label(output_frame, text="No directory selected")
        self.output_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        ttk.Button(output_frame, text="Select Output Directory", command=self._select_output).grid(
            row=0, column=1, padx=5, pady=5
        )

        # Metadata options
        metadata_frame = ttk.LabelFrame(self.root, text="Metadata Options", padding="5")
        metadata_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Publisher name
        ttk.Label(metadata_frame, text="Publisher Name:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(metadata_frame, textvariable=self.publisher_name).grid(
            row=0, column=1, padx=5, pady=5, sticky="ew"
        )

        # Publisher URI
        ttk.Label(metadata_frame, text="Publisher URI:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(metadata_frame, textvariable=self.publisher_uri).grid(
            row=1, column=1, padx=5, pady=5, sticky="ew"
        )

        # License URI
        ttk.Label(metadata_frame, text="License URI:").grid(
            row=2, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(metadata_frame, textvariable=self.license_uri).grid(
            row=2, column=1, padx=5, pady=5, sticky="ew"
        )

        # Base URI (Added)
        ttk.Label(metadata_frame, text="Base URI:").grid(
            row=3, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(metadata_frame, textvariable=self.base_uri).grid(
            row=3, column=1, padx=5, pady=5, sticky="ew"
        )

        # AI Features section
        ai_frame = ttk.LabelFrame(self.root, text="AI Features (Experimental)", padding="5")
        ai_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Enable AI checkbox
        ttk.Checkbutton(ai_frame, text="Enable AI features", variable=self.enable_ai).grid(
            row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w"
        )

        # LLM Provider
        ttk.Label(ai_frame, text="LLM Provider:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        provider_combo = ttk.Combobox(ai_frame, textvariable=self.llm_provider,
                                      values=["openai", "gemini"], state="readonly")
        provider_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # API Key
        ttk.Label(ai_frame, text="API Key:").grid(
            row=2, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(ai_frame, textvariable=self.llm_api_key, show="*").grid(
            row=2, column=1, padx=5, pady=5, sticky="ew"
        )

        # Skip options
        ttk.Checkbutton(ai_frame, text="Skip header generation", variable=self.skip_header_ai).grid(
            row=3, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Checkbutton(ai_frame, text="Skip datatype validation", variable=self.skip_datatype_ai).grid(
            row=3, column=1, padx=5, pady=5, sticky="w"
        )

        # Progress
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Convert button
        self.convert_btn = ttk.Button(
            self.root, text="Convert", command=self._convert, state="disabled"
        )
        self.convert_btn.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

        # Configure grid
        self.root.columnconfigure(0, weight=1)
        for frame in (input_frame, output_frame, metadata_frame, ai_frame):
            frame.columnconfigure(1, weight=1)

    def _select_input(self):
        """Handle input file selection."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file = filename
            self.input_label.config(text=os.path.basename(filename))
            self._update_convert_button()

    def _select_output(self):
        """Handle output directory selection."""
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_dir = dirname
            self.output_label.config(text=os.path.basename(dirname) or dirname)
            self._update_convert_button()

    def _update_convert_button(self):
        """Enable/disable convert button based on selections."""
        if self.input_file and self.output_dir:
            self.convert_btn.config(state="normal")

        else:
            self.convert_btn.config(state="disabled")

    def _convert(self):
        """Handle file conversion."""
        self.progress.start()
        self.convert_btn.config(state="disabled")
        try:
            self.process_file()
        except FileNotFoundError as e:
            messagebox.showerror("Input File Error", str(e))
        except ValueError as e:
            messagebox.showerror("Output Directory Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()
            self.convert_btn.config(state="normal")

    def process_file(self) -> None:
        """Process the selected Excel file."""
        try:
            validate_input_file(self.input_file)
            validate_output_dir(self.output_dir)

            # Read Excel file into memory
            with open(self.input_file, "rb") as f:
                excel_bytes = f.read()            # Process Excel file and get CSV files and metadata
            csv_files, metadata_buffer = process_excel_in_memory(
                excel_bytes,
                os.path.basename(self.input_file), # Pass the original Excel filename
                self.output_dir,
                base_uri=self.base_uri.get(),
                publisher_uri=self.publisher_uri.get(),
                publisher_name=self.publisher_name.get(),
                license_uri=self.license_uri.get(),
                enable_ai=self.enable_ai.get(),
                llm_provider=self.llm_provider.get(),
                llm_api_key=self.llm_api_key.get() if self.llm_api_key.get().strip() else None,
                skip_header_ai=self.skip_header_ai.get(),
                skip_datatype_ai=self.skip_datatype_ai.get()
            )

            # Save metadata
            metadata_format = "turtle"  # Could be made configurable in GUI
            metadata_file = os.path.join(self.output_dir, f"metadata.{metadata_format}")
            with open(metadata_file, "wb") as f:
                f.write(metadata_buffer.getvalue())

            message = f"Successfully processed {len(csv_files)} tables:\n"
            message += "\n".join(f"- {os.path.basename(f)}" for f in csv_files)
            message += f"\n\nMetadata saved to: {os.path.basename(metadata_file)}"

            messagebox.showinfo("Success", message)

        except Exception as e:
            messagebox.showerror("Error", str(e))

def main():
    """Start the GUI application."""
    root = tk.Tk()
    app = ExcelConverterGUI(root)
    root.mainloop()
