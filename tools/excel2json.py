import json
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

import pandas as pd

class Excel2JsonTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        file_meta = tool_parameters['file']
        try:
            xls = pd.ExcelFile(file_meta.url)
            sheet_names = xls.sheet_names

            if len(sheet_names) > 1:
                # Multiple sheets
                all_sheets_data = pd.read_excel(file_meta.url, sheet_name=None, dtype=str)
                json_output = {}
                for sheet_name, df in all_sheets_data.items():
                    # Convert DataFrame to a list of records (dicts)
                    json_output[sheet_name] = json.loads(df.to_json(orient="records", force_ascii=False))
                
                # Convert the entire structure to a JSON string
                yield self.create_text_message(json.dumps(json_output, ensure_ascii=False, indent=2))
            else:
                # Single sheet
                df = pd.read_excel(file_meta.url, dtype=str)
                yield self.create_text_message(df.to_json(orient="records", force_ascii=False))

        except Exception as e:
            raise Exception(f"Error processing Excel file: {str(e)}")
