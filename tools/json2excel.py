import json
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

import pandas as pd
from io import BytesIO

class Json2ExcelTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        json_str = tool_parameters['json_str']
        
        try:
            data = json.loads(json_str)
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format: {e}")

        excel_buffer = BytesIO()
        try:
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                if isinstance(data, dict):
                    # If the top level is an object, treat keys as sheet names
                    for sheet_name, records in data.items():
                        if not isinstance(records, list):
                            raise Exception(f"Value for sheet '{sheet_name}' must be a list of records.")
                        df = pd.DataFrame(records)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                elif isinstance(data, list):
                    # If the top level is an array, write to a single sheet
                    df = pd.DataFrame(data)
                    df.to_excel(writer, sheet_name='Sheet1', index=False)
                else:
                    raise Exception("JSON must be an object (for multiple sheets) or an array (for a single sheet).")
        except Exception as e:
            raise Exception(f"Error converting data to Excel: {str(e)}")

        # create a blob with the excel bytes
        try:
            excel_buffer.seek(0)
            excel_bytes = excel_buffer.getvalue()
            filename = tool_parameters.get('filename', 'Converted Data')
            filename = f"{filename.replace(' ', '_')}.xlsx"

            yield self.create_text_message(f"Excel file '{filename}' generated successfully")

            yield self.create_blob_message(
                    blob=excel_bytes,
                    meta={
                        "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "filename": filename
                    }
                )
        except Exception as e:
            raise Exception(f"Error creating Excel file message: {str(e)}")
