"""
Module: excel_email_sender

This module provides functionality to send an email via Outlook with
content copied directly from a specified Excel sheet range.

It uses the `pywin32` library to interact with Microsoft Excel and Outlook.
"""

import os
import time
import logging
from dataclasses import dataclass
from typing import Optional
import win32com.client as win32
from win32com.client import Dispatch

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)


@dataclass
class EmailConfig:
    """Configuration for sending the Excel content via Outlook."""
    excel_path: str
    excel_filename: str
    sheet_index: int = 1
    cell_range: str = "A1:B3"
    recipient: str = ""
    subject: str = "Excel Data"
    body_html: str = "Hi There"


class ExcelEmailSender:
    """
    Handles sending an email with Excel range content pasted into the body.

    Follows the Single Responsibility Principle (SRP):
    This class is responsible only for reading Excel data and sending it via Outlook.
    """

    def __init__(self, config: EmailConfig) -> None:
        self.config = config
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.outlook = Dispatch("Outlook.Application")

    def send_email(self) -> None:
        """
        Sends the email with the Excel range content embedded.
        Raises:
            FileNotFoundError: If the Excel file does not exist.
            Exception: If there is an issue with Outlook or Excel automation.
        """
        excel_file = os.path.join(self.config.excel_path, self.config.excel_filename)

        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"Excel file not found at: {excel_file}")

        try:
            logging.info("Opening Excel workbook...")
            book = self.excel.Workbooks.Open(excel_file)
            sheet = book.Sheets(self.config.sheet_index)

            logging.info("Copying Excel range...")
            sheet.Range(self.config.cell_range).Copy()

            logging.info("Creating Outlook email draft...")
            msg = self.outlook.CreateItem(0)  # 0 = Mail Item
            msg.HTMLBody = self.config.body_html

            # Display first so WordEditor is ready
            msg.Display()
            time.sleep(0.5)

            logging.info("Pasting Excel content into the email body...")
            msg.GetInspector.WordEditor.Range(Start=0, End=0).Paste()

            msg.To = self.config.recipient
            msg.Subject = self.config.subject
            msg.BodyFormat = 2  # 2 = HTML format

            logging.info("Sending email...")
            msg.Send()

            logging.info("Email sent successfully.")

        except Exception as e:
            logging.error(f"Failed to send email: {e}")
            raise

        finally:
            logging.info("Closing Excel workbook...")
            if 'book' in locals():
                book.Close(SaveChanges=False)

    def __del__(self):
        """Ensures Excel application is properly quit."""
        try:
            self.excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    config = EmailConfig(
        excel_path=r"D:\\",
        excel_filename="Book1.xlsx",
        sheet_index=1,
        cell_range="A1:B3",
        recipient="myemail@live.com",
        subject="Excel Data",
        body_html="Hi There, Good Day!"
    )

    sender = ExcelEmailSender(config)
    sender.send_email()
