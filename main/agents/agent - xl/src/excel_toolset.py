import logging
import os
import json
import pandas as pd
from typing import Any, List, Dict, Union

logger = logging.getLogger(__name__)


class ExcelToolset:
    """Toolset for generating Excel spreadsheets."""

    def __init__(self, host: str, port: int):
        self.host = host
        self.port = port
        self.output_dir = "outputs"
        os.makedirs(self.output_dir, exist_ok=True)
        logger.info(
            f"Initialized ExcelToolset. Saving files to ./{self.output_dir}")

    # THE FIX: Made async and returning a clean text string
    async def generate_excel(self, filename: str, sheet_name: str, data: Union[List[Dict[str, Any]], str]) -> str:
        """
        Generates an Excel (.xlsx) file from structured tabular data.

        Args:
            filename: Output filename (e.g., budget.xlsx)
            sheet_name: Name of the sheet (e.g., Q1_Data)
            data: A list of dictionaries where keys are column headers and values are row data.
        """
        try:
            # Handle stringified JSON from LLM
            if isinstance(data, str):
                logger.info("Data received as a string. Parsing JSON...")
                data = json.loads(data)

            logger.info(
                f"Generating Excel file: {filename} with {len(data)} rows.")

            if not filename.endswith('.xlsx'):
                filename += '.xlsx'

            filepath = os.path.join(self.output_dir, filename)

            # Generate the Excel file
            df = pd.DataFrame(data)
            df.to_excel(filepath, index=False, sheet_name=sheet_name)

            download_url = f"http://{self.host}:{self.port}/outputs/{filename}"

            # THE FIX: Return standard Markdown for the Nasiko UI
            return f"✅ Successfully generated **{filename}**!\n\n📊 [Download your spreadsheet here]({download_url})"

        except Exception as e:
            logger.error(f"Error generating Excel file: {e}")
            return f"❌ Failed to generate spreadsheet: {str(e)}"

    def get_tools(self) -> dict[str, Any]:
        return {
            "generate_excel": self,
        }
