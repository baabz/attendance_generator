"""
Attendance Checker Core Package
"""
from .models import Student, VenueAllocation, ExamConfig
from .parser import (
    parse_excel,
    parse_csv,
    parse_docx,
    parse_raw_text,
    generate_blank_excel_template,
    detect_columns
)
from .shuffler import reshuffle_students, allocate_to_venues
from .docx_generator import build_attendance_docx
from .pdf_generator import build_attendance_pdf
from .door_notices import build_door_notices_docx, build_door_notices_pdf
from .envelope_tags import build_envelope_tags_docx, build_envelope_tags_pdf

__all__ = [
    'Student',
    'VenueAllocation',
    'ExamConfig',
    'parse_excel',
    'parse_csv',
    'parse_docx',
    'parse_raw_text',
    'generate_blank_excel_template',
    'detect_columns',
    'reshuffle_students',
    'allocate_to_venues',
    'build_attendance_docx',
    'build_attendance_pdf',
    'build_door_notices_docx',
    'build_door_notices_pdf',
    'build_envelope_tags_docx',
    'build_envelope_tags_pdf'
]
