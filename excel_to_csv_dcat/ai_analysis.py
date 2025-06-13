"""AI-powered analysis for CSV header generation and datatype validation."""

import os
import json
from typing import List, Optional, Dict, Union
from io import StringIO
import pandas as pd


class LLMAnalyzer:
    """Base class for LLM-powered analysis."""
    
    def __init__(self, api_key: str):
        self.api_key = api_key
    
    def suggest_csv_headers(self, csv_sample_data: str, table_title: Optional[str] = None) -> List[str]:
        """Suggest appropriate headers for a headerless CSV."""
        raise NotImplementedError
    
    def suggest_column_datatypes(self, column_info: Dict[str, str]) -> Dict[str, str]:
        """Suggest appropriate datatypes for CSV columns."""
        raise NotImplementedError


class OpenAIAnalyzer(LLMAnalyzer):
    """OpenAI GPT-powered analyzer."""
    
    def __init__(self, api_key: str):
        super().__init__(api_key)
        try:
            import openai
            self.client = openai.OpenAI(api_key=api_key)
        except ImportError:
            raise ImportError("OpenAI library not installed. Run: pip install openai")
    
    def suggest_csv_headers(self, csv_sample_data: str, table_title: Optional[str] = None) -> List[str]:
        """Suggest appropriate headers for a headerless CSV using OpenAI GPT."""
        # Prepare the prompt
        title_context = f"The table title is: '{table_title}'. " if table_title else ""
        
        prompt = f"""
You are an expert data analyst. {title_context}Below is a sample of a CSV file without headers. 
Please suggest appropriate column headers based on the data content.

CSV Data Sample:
{csv_sample_data}

Requirements:
- Provide concise, descriptive column names
- Use snake_case format (e.g., "first_name", "total_amount")
- Return only a JSON array of header names
- The number of headers should match the number of columns in the data
- Base suggestions on the actual data patterns you observe

Example response format: ["column_1", "column_2", "column_3"]
"""
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful data analysis assistant that suggests column headers for CSV files."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=200,
                temperature=0,
                timeout=30  # 30 second timeout
            )
            
            # Parse the response
            content = response.choices[0].message.content.strip()
            
            # Extract JSON array from response
            try:
                start_idx = content.find('[')
                end_idx = content.rfind(']') + 1
                if start_idx >= 0 and end_idx > start_idx:
                    json_str = content[start_idx:end_idx]
                    headers = json.loads(json_str)
                    return headers if isinstance(headers, list) else []
                else:
                    return []
            except json.JSONDecodeError:
                return []                
        except Exception as e:
                print(f"Warning: OpenAI API call failed for header suggestion: {e}")
        return []
    
    def suggest_column_datatypes(self, column_info: Dict[str, str]) -> Dict[str, str]:
        """Suggest appropriate datatypes for CSV columns using OpenAI GPT."""
        prompt = f"""
You are an expert data analyst. Below is information about CSV columns with their names, current inferred types, and sample data.
Please suggest the most appropriate XSD datatype for each column.

Column Information:
{json.dumps(column_info, indent=2)}

Available XSD datatypes:
- xsd:string (for text data)
- xsd:integer (for whole numbers)
- xsd:decimal (for decimal numbers)
- xsd:boolean (for true/false values)
- xsd:date (for dates in YYYY-MM-DD format)
- xsd:dateTime (for datetime values)
- xsd:time (for time values)

Requirements:
- Return only a JSON object mapping column names to XSD datatype strings
- Consider the data patterns and choose the most specific appropriate type
- Default to xsd:string when uncertain

Example response format: {{"column_1": "xsd:integer", "column_2": "xsd:string"}}
"""
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful data analysis assistant that suggests appropriate datatypes for CSV columns."},
                    {"role": "user", "content": prompt}                ],
                max_tokens=300,
                temperature=0,
                timeout=30  # 30 second timeout
            )
            
            # Parse the response
            content = response.choices[0].message.content.strip()
            
            try:
                start_idx = content.find('{')
                end_idx = content.rfind('}') + 1
                if start_idx >= 0 and end_idx > start_idx:
                    json_str = content[start_idx:end_idx]
                    datatypes = json.loads(json_str)
                    return datatypes if isinstance(datatypes, dict) else {}
                else:
                    return {}
            except json.JSONDecodeError:
                return {}
                
        except Exception as e:
            return {}


class GeminiAnalyzer(LLMAnalyzer):
    """Google Gemini-powered analyzer."""
    
    def __init__(self, api_key: str):
        super().__init__(api_key)
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel('gemini-1.5-flash')
        except ImportError:
            raise ImportError("Google Generative AI library not installed. Run: pip install google-generativeai")
    
    def suggest_csv_headers(self, csv_sample_data: str, table_title: Optional[str] = None) -> List[str]:
        """Suggest appropriate headers for a headerless CSV using Google Gemini."""
        # Prepare the prompt
        title_context = f"The table title is: '{table_title}'. " if table_title else ""
        
        prompt = f"""
You are an expert data analyst. {title_context}Below is a sample of a CSV file without headers. 
Please suggest appropriate column headers based on the data content.

CSV Data Sample:
{csv_sample_data}

Requirements:
- Provide concise, descriptive column names
- Use snake_case format (e.g., "first_name", "total_amount")
- Return only a JSON array of header names
- The number of headers should match the number of columns in the data
- Base suggestions on the actual data patterns you observe

Example response format: ["column_1", "column_2", "column_3"]
"""
        
        try:
            # Add timeout and generation config for more reliable responses
            import google.generativeai as genai
            generation_config = genai.types.GenerationConfig(
                temperature=0,
                top_p=1,
                top_k=1,
                max_output_tokens=200,
            )
            
            response = self.model.generate_content(
                prompt,
                generation_config=generation_config,
                request_options={'timeout': 30}  # 30 second timeout
            )
            content = response.text.strip()
            
            # Extract JSON array from response
            try:
                start_idx = content.find('[')
                end_idx = content.rfind(']') + 1
                if start_idx >= 0 and end_idx > start_idx:
                    json_str = content[start_idx:end_idx]
                    headers = json.loads(json_str)
                    return headers if isinstance(headers, list) else []
                else:
                    return []
            except json.JSONDecodeError:
                return []                
        except Exception as e:
            print(f"Warning: Gemini API call failed for header suggestion: {e}")
            return []
    
    def suggest_column_datatypes(self, column_info: Dict[str, str]) -> Dict[str, str]:
        """Suggest appropriate datatypes for CSV columns using Google Gemini."""
        prompt = f"""
You are an expert data analyst. Below is information about CSV columns with their names, current inferred types, and sample data.
Please suggest the most appropriate XSD datatype for each column.

Column Information:
{json.dumps(column_info, indent=2)}

Available XSD datatypes:
- xsd:string (for text data)
- xsd:integer (for whole numbers)
- xsd:decimal (for decimal numbers)
- xsd:boolean (for true/false values)
- xsd:date (for dates in YYYY-MM-DD format)
- xsd:dateTime (for datetime values)
- xsd:time (for time values)

Requirements:
- Return only a JSON object mapping column names to XSD datatype strings
- Consider the data patterns and choose the most specific appropriate type
- Default to xsd:string when uncertain

Example response format: {{"column_1": "xsd:integer", "column_2": "xsd:string"}}
"""
        
        try:
            # Add timeout and generation config for more reliable responses
            import google.generativeai as genai
            generation_config = genai.types.GenerationConfig(
                temperature=0,
                top_p=1,
                top_k=1,
                max_output_tokens=300,
            )
            
            response = self.model.generate_content(
                prompt,
                generation_config=generation_config,
                request_options={'timeout': 30}  # 30 second timeout
            )
            content = response.text.strip()
            
            try:
                start_idx = content.find('{')
                end_idx = content.rfind('}') + 1
                if start_idx >= 0 and end_idx > start_idx:
                    json_str = content[start_idx:end_idx]
                    datatypes = json.loads(json_str)
                    return datatypes if isinstance(datatypes, dict) else {}
                else:
                    return {}
            except json.JSONDecodeError:
                return {}
                
        except Exception as e:
            print(f"Warning: Gemini API call failed for datatype suggestion: {e}")
            return {}


def get_llm_analyzer(provider: str, api_key: str) -> Union[OpenAIAnalyzer, GeminiAnalyzer]:
    """Factory function to create the appropriate LLM analyzer."""
    if provider.lower() == "openai":
        return OpenAIAnalyzer(api_key)
    elif provider.lower() == "gemini":
        return GeminiAnalyzer(api_key)
    else:
        raise ValueError(f"Unsupported LLM provider: {provider}. Supported providers: 'openai', 'gemini'")


def prepare_csv_sample(csv_file_path: str, max_rows: int = 10) -> str:
    """Prepare a sample of CSV data for LLM analysis."""
    try:
        # Read the CSV file without headers
        df = pd.read_csv(csv_file_path, header=None, nrows=max_rows)
        
        # Convert to string representation
        csv_buffer = StringIO()
        df.to_csv(csv_buffer, index=False, header=False)
        return csv_buffer.getvalue().strip()
    except Exception as e:
        return ""


def prepare_csv_sample_from_content(csv_content: bytes, max_rows: int = 15) -> str:
    """Prepare a sample of CSV data from byte content for LLM analysis."""
    try:
        from io import BytesIO, StringIO
        # Read CSV from bytes
        csv_buffer = BytesIO(csv_content)
        df = pd.read_csv(csv_buffer, header=None, nrows=max_rows)
        
        # Convert to string representation
        output_buffer = StringIO()
        df.to_csv(output_buffer, index=False, header=False)
        return output_buffer.getvalue().strip()
    except Exception as e:
        return ""


def prepare_column_info_for_datatype_analysis(csv_file_path: str, headers: List[str], max_sample_rows: int = 20) -> Dict[str, str]:
    """Prepare column information for datatype analysis from file path."""
    try:
        # Read CSV with the provided headers
        df = pd.read_csv(csv_file_path, names=headers, nrows=max_sample_rows)
        
        column_info = {}
        for col_name in headers:
            if col_name in df.columns:
                # Get sample values (non-null)
                sample_values = df[col_name].dropna().head(5).tolist()
                sample_str = ", ".join([str(val) for val in sample_values])
                
                # Get inferred pandas dtype
                inferred_dtype = str(df[col_name].dtype)
                
                column_info[col_name] = f"Current type: {inferred_dtype}, Sample values: [{sample_str}]"
        
        return column_info
    except Exception as e:
        return {}


def prepare_column_info_in_memory(csv_content: bytes, headers: List[str], max_sample_rows: int = 20) -> Dict[str, str]:
    """Prepare column information for datatype analysis from in-memory CSV content."""
    try:
        from io import BytesIO
        
        # Read CSV from bytes with provided headers
        csv_buffer = BytesIO(csv_content)
        df = pd.read_csv(csv_buffer, names=headers, nrows=max_sample_rows)
        
        column_info = {}
        for col_name in headers:
            if col_name in df.columns:
                # Get sample values (non-null)
                sample_values = df[col_name].dropna().head(5).tolist()
                sample_str = ", ".join([str(val) for val in sample_values])
                
                # Get inferred pandas dtype
                inferred_dtype = str(df[col_name].dtype)
                
                column_info[col_name] = f"Current type: {inferred_dtype}, Sample values: [{sample_str}]"
        
        return column_info
    except Exception as e:
        return {}
