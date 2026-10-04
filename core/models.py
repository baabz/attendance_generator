"""
Data models for Exam Attendance Sheet and Hall Allocation.
"""
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Student:
    reg_no: str
    name: str
    original_sn: Optional[int] = None

    def __post_init__(self):
        self.reg_no = str(self.reg_no).strip().upper()
        self.name = str(self.name).strip().upper()


@dataclass
class VenueAllocation:
    venue_name: str
    capacity: int
    students: List[Student] = field(default_factory=list)


@dataclass
class ExamConfig:
    session_semester: str = "2nd/4th SEMESTER, 2025/2026 SESSION (REGULAR)"
    department: str = ""
    college: str = "KADUNA POLYTECHNIC"
    course_title: str = ""
    course_code: str = ""
    exam_date: str = ""
    time_in: str = ""
    time_out: str = ""
    level: str = ""
    chief_invigilator: str = "______________________________________"
    invigilator_count: int = 3
    continuous_sn: bool = False
