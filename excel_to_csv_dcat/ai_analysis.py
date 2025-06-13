"""Unified AI-powered analysis for CSV header generation and datatype validation."""

import os
import json
from typing import List, Optional, Dict, Union
from io import StringIO
import pandas as pd
from tqdm import tqdm


class UnifiedLLMAnalyzer:
    """Unified LLM-powered analyzer supporting both OpenAI and Gemini."""
    
    def __init__(self, provider: str, api_key: str):
        self.provider = provider.lower()
        self.api_key = api_key
        self.client = None  # For OpenAI
        self.model = None   # For Gemini
        self.genai = None   # For Gemini module        
        if self.provider not in ['openai', 'gemini']:
            raise ValueError(f"Unsupported provider: {provider}. Supported: 'openai', 'gemini'")
    
    def _ensure_client_ready(self):
        """Ensure the appropriate client is initialized based on provider."""
        if self.provider == 'openai' and self.client is None:
            try:
                import openai
                self.client = openai.OpenAI(api_key=self.api_key)
            except ImportError:
                raise ImportError("OpenAI library not installed. Run: pip install openai")
        
        elif self.provider == 'gemini' and self.model is None:
            try:
                import google.generativeai as genai
                self.genai = genai  # Store for later use
                genai.configure(api_key=self.api_key)
                self.model = genai.GenerativeModel('gemini-2.0-flash-lite')
            except ImportError:
                raise ImportError("Google Generative AI library not installed. Run: pip install google-generativeai")
    
    @staticmethod
    def _parse_json_response(content: str, expected_type: type, fallback_value=None):
        """Parse JSON response from AI model with robust error handling."""
        try:
            if expected_type == list:
                start_char, end_char = '[', ']'
            elif expected_type == dict:
                start_char, end_char = '{', '}'
            else:
                raise ValueError(f"Unsupported expected_type: {expected_type}")
            
            start_idx = content.find(start_char)
            end_idx = content.rfind(end_char) + 1
            
            if start_idx >= 0 and end_idx > start_idx:
                json_str = content[start_idx:end_idx]
                parsed_data = json.loads(json_str)
                
                if isinstance(parsed_data, expected_type):
                    return parsed_data
            
            return fallback_value or ([] if expected_type == list else {})
            
        except json.JSONDecodeError:
            return fallback_value or ([] if expected_type == list else {})
    
    @staticmethod
    def _build_batch_header_prompt(tables_data: List[Dict[str, str]]) -> str:
        """Build consistent batch prompt for header suggestions."""
        batch_prompt = """You are an expert data analyst. Below are multiple CSV tables without headers from the same Excel file.
For each table, suggest appropriate column headers based on the data content.

Tables:
"""
        
        for i, table_data in enumerate(tables_data, 1):
            table_title = table_data.get('table_title', 'Unknown')
            csv_sample = table_data['csv_sample_data']
            batch_prompt += f"""
Table {i} (Title: {table_title}):
{csv_sample}

"""
        
        batch_prompt += """
Requirements:
- Provide concise, descriptive column names for each table
- Use snake_case format (e.g., "first_name", "total_amount")
- Return only a JSON object where keys are "table_1", "table_2", etc. and values are arrays of header names
- The number of headers should match the number of columns in each table's data
- Base suggestions on the actual data patterns you observe

Example response format: {"table_1": ["col1", "col2"], "table_2": ["header1", "header2", "header3"]}
"""
        return batch_prompt
    
    @staticmethod
    def _build_batch_datatype_prompt(tables_column_info: Dict[str, Dict[str, str]]) -> str:
        """Build consistent batch prompt for datatype suggestions."""
        batch_prompt = """You are an expert data analyst. Below is information about CSV columns from multiple tables in the same Excel file.
For each table, suggest the most appropriate XSD datatype for each column.

Tables:
"""
        
        for i, (table_name, column_info) in enumerate(tables_column_info.items(), 1):
            batch_prompt += f"""
Table {i} ({table_name}):
{json.dumps(column_info, indent=2)}

"""
        
        batch_prompt += """Available XSD datatypes:
- xsd:string (for text data)
- xsd:integer (for whole numbers)
- xsd:decimal (for decimal numbers)
- xsd:boolean (for true/false values)
- xsd:date (for dates in YYYY-MM-DD format)
- xsd:dateTime (for datetime values)
- xsd:time (for time values)

Requirements:
- Return only a JSON object where keys are "table_1", "table_2", etc.
- Each table's value should be a JSON object mapping column names to XSD datatype strings
- Consider the data patterns and choose the most specific appropriate type
- Default to xsd:string when uncertain

Example response format: {"table_1": {"col1": "xsd:integer", "col2": "xsd:string"}, "table_2": {"header1": "xsd:decimal"}}
"""
        return batch_prompt
    
    def _call_ai_api(self, prompt: str, timeout: int = 30) -> str:
        """Make API call to the configured provider."""
        self._ensure_client_ready()
        
        if self.provider == 'openai':
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful data analysis assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0,
                timeout=timeout
            )
            return response.choices[0].message.content.strip()        
        elif self.provider == 'gemini':
            generation_config = self.genai.types.GenerationConfig(
                temperature=0,
                top_p=1,
                top_k=1,
            )
            response = self.model.generate_content(
                prompt,
                generation_config=generation_config,
                request_options={'timeout': timeout}            )
            return response.text.strip()
    
    def suggest_csv_headers_batch(self, tables_data: List[Dict[str, str]]) -> Dict[str, List[str]]:
        """Suggest appropriate headers for multiple tables in a single API call."""
        if not tables_data:
            return {}
            
        with tqdm(total=1, desc="AI header generation", unit="batch", leave=False) as pbar:
            batch_prompt = self._build_batch_header_prompt(tables_data)
            
            try:
                content = self._call_ai_api(batch_prompt, timeout=60)
                batch_headers = self._parse_json_response(content, dict, {})
                
                # Map back to table names
                result = {}
                for i, table_data in enumerate(tables_data, 1):
                    table_key = f"table_{i}"
                    table_name = table_data['table_name']
                    if table_key in batch_headers and isinstance(batch_headers[table_key], list):
                        result[table_name] = batch_headers[table_key]
                    else:
                        result[table_name] = []                
                pbar.update(1)
                return result
                    
            except Exception as e:
                pbar.update(1)
                return {table_data['table_name']: [] for table_data in tables_data}
    
    def suggest_column_datatypes_batch(self, tables_column_info: Dict[str, Dict[str, str]]) -> Dict[str, Dict[str, str]]:
        """Suggest appropriate datatypes for multiple tables in a single API call."""
        if not tables_column_info:
            return {}
            
        with tqdm(total=1, desc="AI datatype validation", unit="batch", leave=False) as pbar:
            batch_prompt = self._build_batch_datatype_prompt(tables_column_info)
            
            try:
                content = self._call_ai_api(batch_prompt, timeout=60)
                batch_datatypes = self._parse_json_response(content, dict, {})
                
                # Map back to table names
                result = {}
                table_names = list(tables_column_info.keys())
                for i, table_name in enumerate(table_names, 1):
                    table_key = f"table_{i}"
                    if table_key in batch_datatypes and isinstance(batch_datatypes[table_key], dict):
                        result[table_name] = batch_datatypes[table_key]
                    else:
                        result[table_name] = {}
                
                pbar.update(1)
                return result
                    
            except Exception as e:
                pbar.update(1)
                return {table_name: {} for table_name in tables_column_info.keys()}


def get_llm_analyzer(provider: str, api_key: str) -> UnifiedLLMAnalyzer:
    """Factory function to create the unified LLM analyzer."""
    return UnifiedLLMAnalyzer(provider, api_key)


def prepare_csv_sample_from_content(csv_content: bytes, max_rows: int = 15) -> str:
    """Prepare a sample of CSV data from byte content for LLM analysis."""
    try:
        from io import BytesIO, StringIO
        csv_buffer = BytesIO(csv_content)
        df = pd.read_csv(csv_buffer, header=None, nrows=max_rows)        
        output_buffer = StringIO()
        df.to_csv(output_buffer, index=False, header=False)
        return output_buffer.getvalue().strip()
    except Exception as e:
        return ""


def prepare_column_info_in_memory(csv_content: bytes, headers: List[str], max_sample_rows: int = 20) -> Dict[str, str]:
    """Prepare column information for datatype analysis from in-memory CSV content."""
    try:
        from io import BytesIO
        csv_buffer = BytesIO(csv_content)
        df = pd.read_csv(csv_buffer, names=headers, nrows=max_sample_rows)
        
        column_info = {}
        for col_name in headers:
            if col_name in df.columns:
                sample_values = df[col_name].dropna().head(5).tolist()
                sample_str = ", ".join([str(val) for val in sample_values])
                inferred_dtype = str(df[col_name].dtype)
                column_info[col_name] = f"Current type: {inferred_dtype}, Sample values: [{sample_str}]"
        
        return column_info
    except Exception as e:
        return {}
