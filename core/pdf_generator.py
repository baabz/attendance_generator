"""
PDF Generator for Official Exam Attendance Sheets using ReportLab.
"""
import io
from typing import List, Union
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

from .models import ExamConfig, VenueAllocation


def build_attendance_pdf(
    config: ExamConfig,
    allocations: List[VenueAllocation],
    output_target: Union[str, io.BytesIO, None] = None
) -> bytes:
    """
    Generate a clean, printable PDF containing all allocated halls.
    """
    buffer = io.BytesIO() if output_target is None or not isinstance(output_target, str) else output_target

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.5 * inch
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'HeaderTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=13,
        alignment=1,  # Center
        spaceAfter=2
    )

    meta_left_style = ParagraphStyle(
        'MetaLeft',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        alignment=0  # Left
    )

    meta_right_style = ParagraphStyle(
        'MetaRight',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        alignment=0  # Left
    )

    invig_style = ParagraphStyle(
        'InvigLine',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        alignment=2  # Right
    )

    tbl_header_style = ParagraphStyle(
        'TblHdr',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=10,
        alignment=1  # Center
    )

    tbl_cell_center = ParagraphStyle(
        'TblCellCenter',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=10,
        alignment=1
    )

    tbl_cell_left = ParagraphStyle(
        'TblCellLeft',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=10,
        alignment=0
    )

    elements = []
    running_sn = 1

    for hall_idx, alloc in enumerate(allocations):
        if hall_idx > 0:
            elements.append(PageBreak())

        # Header Titles
        elements.append(Paragraph(config.session_semester.strip(), title_style))
        elements.append(Paragraph(config.department.strip(), title_style))
        elements.append(Paragraph(config.college.strip(), title_style))
        elements.append(Paragraph("ATTENDANCE SHEET", title_style))
        elements.append(Spacer(1, 4))

        # Metadata Table
        time_in_str = config.time_in.strip() if config.time_in.strip() else "..................................."
        time_out_str = config.time_out.strip() if config.time_out.strip() else "..................................."
        chief_str = config.chief_invigilator.strip() if config.chief_invigilator.strip() else "____________________"

        meta_data = [
            [
                Paragraph(f"COURSE: {config.course_title.strip()}", meta_left_style),
                Paragraph(f"CODE: {config.course_code.strip()}", meta_right_style)
            ],
            [
                Paragraph(f"HALL: {alloc.venue_name.strip()}", meta_left_style),
                Paragraph(f"DATE: {config.exam_date.strip()}", meta_right_style)
            ],
            [
                Paragraph(f"TIME IN: {time_in_str}", meta_left_style),
                Paragraph(f"TIME OUT: {time_out_str}", meta_right_style)
            ],
            [
                Paragraph(f"LEVEL: {config.level.strip()}", meta_left_style),
                Paragraph(f"Chief Invigilator: {chief_str}", meta_right_style)
            ]
        ]

        meta_table = Table(meta_data, colWidths=[4.2 * inch, 3.3 * inch])
        meta_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ]))
        elements.append(meta_table)

        # Invigilators
        elements.append(Spacer(1, 2))
        for num in range(1, (config.invigilator_count or 3) + 1):
            elements.append(Paragraph(f"{num}. .....................................................................", invig_style))

        elements.append(Spacer(1, 6))

        # Student Table
        table_data = [
            [
                Paragraph("S/N", tbl_header_style),
                Paragraph("REG NO", tbl_header_style),
                Paragraph("NAMES", tbl_header_style),
                Paragraph("BOOKLET NO", tbl_header_style),
                Paragraph("SIGN IN", tbl_header_style),
                Paragraph("SIGN OUT", tbl_header_style)
            ]
        ]

        for s_idx, student in enumerate(alloc.students):
            sn_val = str(running_sn) if config.continuous_sn else str(s_idx + 1)
            if config.continuous_sn:
                running_sn += 1

            table_data.append([
                Paragraph(sn_val, tbl_cell_center),
                Paragraph(student.reg_no, tbl_cell_center),
                Paragraph(student.name, tbl_cell_left),
                "",
                "",
                ""
            ])

        col_widths = [0.45 * inch, 1.45 * inch, 2.60 * inch, 1.10 * inch, 0.95 * inch, 0.95 * inch]
        student_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        student_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ]))

        elements.append(student_table)

    doc.build(elements)

    if output_target is None:
        return buffer.getvalue()
    elif isinstance(output_target, str):
        with open(output_target, 'rb') as f:
            return f.read()
    elif hasattr(output_target, 'getvalue'):
        return output_target.getvalue()
    return b''
