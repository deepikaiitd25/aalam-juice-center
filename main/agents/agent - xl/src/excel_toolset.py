import logging
import os
import pandas as pd
from typing import Any, List, Dict
from pydantic import BaseModel

logger = logging.getLogger(__name__)


class ExcelGenerationResponse(BaseModel):
    status: str
    file_url: str | None = None
    error_message: str | None = None


class ExcelToolset:
    """Toolset for generating Excel spreadsheets."""

    def __init__(self, host: str, port: int):
        self.host = host
        self.port = port
        # Create an outputs directory to store the generated files
        self.output_dir = "outputs"
        os.makedirs(self.output_dir, exist_ok=True)
        logger.info(
            f"Initialized ExcelToolset. Saving files to ./{self.output_dir}")

    def generate_excel(self, filename: str, sheet_name: str, data: List[Dict[str, Any]]) -> ExcelGenerationResponse:
        """
        Generates an Excel (.xlsx) file from structured tabular data.

        Args:
            filename: The name of the file (e.g., 'sales_tracker.xlsx'). Must end in .xlsx.
            sheet_name: The name of the primary worksheet.
            data: A list of dictionaries representing the rows of the spreadsheet. Keys are column headers.

        Returns:
            ExcelGenerationResponse: Contains the status and the download URL of the generated file.
        """
        try:
            logger.info(
                f"Generating Excel file: {filename} with {len(data)} rows.")

            # Ensure correct extension
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'

            filepath = os.path.join(self.output_dir, filename)

            # Convert the list of dicts from the LLM into a Pandas DataFrame
            df = pd.DataFrame(data)

            # Write to Excel
            df.to_excel(filepath, index=False, sheet_name=sheet_name)

            # Construct a download URL (we will mount this directory in Starlette in __main__.py)
            download_url = f"http://{self.host}:{self.port}/outputs/{filename}"

            return ExcelGenerationResponse(
                status="success",
                file_url=download_url,
            )

        except Exception as e:
            logger.error(f"Error generating Excel file: {e}")
            return ExcelGenerationResponse(
                status="error",
                error_message=f"Failed to generate file: {str(e)}",
            )

    def get_tools(self) -> dict[str, Any]:
        """Return dictionary of available tools for OpenAI function calling"""
        return {
            "generate_excel": self,
        }
