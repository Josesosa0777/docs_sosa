import logging
import os
import time
from typing import List, Dict

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    PageTemplate,
    Spacer,
    Frame,
)
from reportlab.pdfgen.canvas import Canvas

from enums import ValidationType


class PDFReportGenerator:
    """Class for PDF Reports."""

    def __init__(self) -> None:
        """
        Initialize the PDF report generator with logging configuration.
        """
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)

    def add_header_footer(self, canvas: Canvas, doc: SimpleDocTemplate, validation_type: ValidationType, part_identifier: str) -> None:
        """
        Add a header and footer to the PDF document.

        Args:
            canvas (Canvas): The canvas to draw on.
            doc (SimpleDocTemplate): The document being created.
            validation_type (ValidationType): Type of validation being performed.
            part_identifier (str): Identifier for the part being reported.

        Returns:
            None
        """
        canvas.saveState()
        # Header
        header_text = f"{validation_type.value} VALIDATION ({part_identifier})"
        canvas.setFont("Helvetica-Bold", 14)
        text_width = canvas.stringWidth(header_text, "Helvetica-Bold", 14)
        canvas.drawCentredString(doc.pagesize[0] / 2, doc.pagesize[1] - 80, header_text)

        # Footer
        footer_text = f"Page {doc.page}"
        canvas.setFont("Helvetica", 10)
        text_width = canvas.stringWidth(footer_text, "Helvetica", 10)
        canvas.drawString((doc.pagesize[0] - text_width) / 2, 40, footer_text)
        canvas.restoreState()

    def generate_pdf(
        self, dfs: List[Dict[str, pd.DataFrame]], validation_type: ValidationType, part_identifier: str
    ) -> bool:
        """
        Generate a PDF report from the given DataFrames.

        Args:
            dfs (List[Dict[str, pd.DataFrame]]): List of dictionaries containing DataFrames and corresponding messages.
            validation_type (ValidationType): Type of validation being performed.
            part_identifier (str): Identifier for the part being reported.

        Returns:
            bool: True if the PDF was successfully created, otherwise False.
        """
        if not dfs or not validation_type:
            return False
        
        pdf_filename = f"{validation_type.value}_Report_{time.strftime('%Y%m%d%H%M%S')}.pdf"
        reports_folder = os.path.join(os.path.abspath("."), "Reports")
        os.makedirs(reports_folder, exist_ok=True)
        pdf_path = os.path.join(reports_folder, pdf_filename)

        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=letter,
            topMargin=1.25 * inch,
            bottomMargin=1.0 * inch,
            leftMargin=0.5 * inch,
            rightMargin=0.5 * inch,
        )

        styles = getSampleStyleSheet()
        content = []

        for element in dfs:
            df = element["df"]
            msg = element["msg"]
            cell_style = ParagraphStyle(
                name="CellStyle",
                parent=styles["Normal"],
                fontName="Times-Roman",
                fontSize=10,
                leading=12,
            )
            paragraph_style = ParagraphStyle(
                name="CustomStyle",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=12,
                leading=14,
                alignment=TA_JUSTIFY,
                spaceAfter=12,
            )

            results_text_html = msg.replace("\n", "<br />")
            results_paragraph = Paragraph(results_text_html, paragraph_style)
            content.append(results_paragraph)

            if not df.empty:
                column_headers = df.columns.tolist()
                table_data = [column_headers]
                if validation_type == ValidationType.ATTRIBUTES:
                    table_data = [column_headers[:-1]]  # Hide last column
                    for row in df.itertuples(index=False):
                        table_data.append(
                            [Paragraph(str(cell), cell_style) for cell in row[:-1]]
                        )
                else:
                    for row in df.itertuples(index=False):
                        table_data.append(
                            [Paragraph(str(cell), cell_style) for cell in row]
                        )

                style = TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("FONTSIZE", (0, 0), (-1, 0), 12),
                        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                        ("GRID", (0, 0), (-1, -1), 1, colors.black),
                        ("FONTNAME", (0, 1), (-1, -1), "Times-Roman"),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ]
                )

                available_width = doc.width - 0.20 * inch
                num_columns = len(table_data[0])
                column_width = available_width / num_columns
                table = Table(
                    table_data, colWidths=[column_width] * num_columns, repeatRows=1
                )
                table.setStyle(style)

                if validation_type in [ValidationType.INI, ValidationType.BOM, ValidationType.ATTRIBUTES, ValidationType.SCHEDULE]:
                    self._apply_conditional_styling(
                        table, df, validation_type, column_headers
                    )

                content.append(table)
                content.append(Spacer(1, 0.5 * inch))
        
        frame = Frame(
            doc.leftMargin,
            doc.bottomMargin,
            doc.width,
            doc.pagesize[1] - doc.bottomMargin - doc.topMargin,
            id="normal",
        )
        template = PageTemplate(
            id="test",
            frames=frame,
            onPage=lambda canvas, doc: self.add_header_footer(
                canvas, doc, validation_type, part_identifier
            ),
        )
        doc.addPageTemplates([template])

        doc.build(
            content,
            onFirstPage=lambda canvas, doc: self.add_header_footer(
                canvas, doc, validation_type, part_identifier
            ),
            onLaterPages=lambda canvas, doc: self.add_header_footer(
                canvas, doc, validation_type, part_identifier
            ),
        )
        os.startfile(pdf_path)
        return True

    def _apply_conditional_styling(self, table: Table, df: pd.DataFrame, validation_type: ValidationType, column_headers: List[str]) -> None:
        """
        Apply conditional styling to the table based on validation type.

        Args:
            table (Table): The table to style.
            df (pd.DataFrame): DataFrame with data to be styled.
            validation_type (ValidationType): Type of validation.
            column_headers (List[str]): List of column headers for the DataFrame.

        Returns:
            None
        """
        conditional_cell_style = ParagraphStyle(
            name="ConditionalCellStyle",
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            textColor=colors.white,
            bold=1,
        )
        table_data = table._cellvalues
        for i in range(1, len(df) + 1):
            if validation_type in [ValidationType.INI, ValidationType.BOM]:
                actual_value = df.loc[i - 1, column_headers[-1]]
                expected_value = df.loc[i - 1, column_headers[-2]]
                if expected_value != actual_value:
                    table.setStyle(
                        TableStyle([("BACKGROUND", (0, i), (-1, i), colors.red)])
                    )
                    table_data[i] = [
                        Paragraph(str(cell.text), conditional_cell_style)
                        for cell in table_data[i]
                    ]
            elif validation_type == ValidationType.ATTRIBUTES:
                param = df.loc[i - 1, column_headers[0]]
                should_be_equal = df.loc[i - 1, column_headers[-1]]
                actual_value = df.loc[i - 1, column_headers[-2]]
                expected_value = df.loc[i - 1, column_headers[-3]]
                if param != "Device Code":
                    if (actual_value != expected_value and should_be_equal) or (
                        not should_be_equal and actual_value == ""
                    ):
                        table.setStyle(
                            TableStyle([("BACKGROUND", (0, i), (-1, i), colors.red)])
                        )
                        table_data[i] = [
                            Paragraph(str(cell.text), conditional_cell_style)
                            for cell in table_data[i]
                        ]
                elif (
                    param == "Device Code"
                    and actual_value not in expected_value.split(" or ")
                ):
                    table.setStyle(
                        TableStyle([("BACKGROUND", (0, i), (-1, i), colors.red)])
                    )
                    table_data[i] = [
                        Paragraph(str(cell.text), conditional_cell_style)
                        for cell in table_data[i]
                    ]
            elif validation_type == ValidationType.SCHEDULE:
                comparison = df.loc[i - 1, column_headers[-1]]
                if comparison == "Incorrect":
                    table.setStyle(
                        TableStyle([("BACKGROUND", (0, i), (-1, i), colors.red)])
                    )
                    table_data[i] = [
                        Paragraph(str(cell.text), conditional_cell_style)
                        for cell in table_data[i]
                    ]
        table._cellvalues = table_data
