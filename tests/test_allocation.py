"""
Automated tests for parser, reshuffler, docx generator, and pdf generator.
"""
import io
import unittest
from core.models import Student, ExamConfig, VenueAllocation
from core.parser import (
    parse_raw_text,
    parse_docx,
    generate_blank_excel_template,
    detect_columns
)
from core.shuffler import reshuffle_students, allocate_to_venues
from core.docx_generator import build_attendance_docx
from core.pdf_generator import build_attendance_pdf


class TestAttendanceChecker(unittest.TestCase):

    def setUp(self):
        self.students = [
            Student(reg_no="COE23HND0120", name="KINSMAN ADAVIZE OMANANYI", original_sn=1),
            Student(reg_no="COE23HND0775", name="EMMANUEL ORDUE LUMUNGA", original_sn=2),
            Student(reg_no="COE24HND0002", name="ENOCH TUKURAH ILIYA", original_sn=3),
            Student(reg_no="COE24HND0004", name="ZEENATULAI SANUSI", original_sn=4),
            Student(reg_no="COE24HND0009", name="HAUWAU RAJI MUHAMMAD", original_sn=5),
            Student(reg_no="COE24HND0010", name="DOUGLAS ISHAKU", original_sn=6),
            Student(reg_no="COE24HND0011", name="ABUBAKAR YUSUF", original_sn=7),
            Student(reg_no="COE24HND0012", name="EMMANUEL MATHEW AMOS", original_sn=8),
            Student(reg_no="COE24HND0013", name="DENNIS DANJUMA", original_sn=9),
            Student(reg_no="COE24HND0014", name="ALIYU HAMISU DANLADI", original_sn=10),
        ]
        self.config = ExamConfig(
            session_semester="2nd/4th SEMESTER, 2025/2026 SESSION (REGULAR)",
            department="DEPARTMENT OF COMPUTER ENGINEERING",
            college="COLLEGE OF ENGINEERING, KADUNA POLYTECHNIC",
            course_title="ARTIFICIAL INTELLIGENCE",
            course_code="CTE 442",
            exam_date="29th Sept, 2026",
            time_in="09:00 AM",
            time_out="12:00 PM",
            level="HND II"
        )

    def test_custom_seed_reshuffle(self):
        # 1. Custom seed must be deterministic
        s1 = reshuffle_students(self.students, seed="DAY1_CTE442")
        s2 = reshuffle_students(self.students, seed="DAY1_CTE442")
        self.assertEqual([s.reg_no for s in s1], [s.reg_no for s in s2])

        # 2. Different seeds must produce different orders
        s3 = reshuffle_students(self.students, seed="DAY2_CTE442")
        self.assertNotEqual([s.reg_no for s in s1], [s.reg_no for s in s3])

        # 3. No loss or duplication of students
        self.assertEqual(len(s1), len(self.students))
        self.assertEqual(set(s.reg_no for s in s1), set(s.reg_no for s in self.students))

    def test_venue_allocation(self):
        venues = [
            ("MINING THEATER-MPRE", 6),
            ("CIVIL ROOM 101/102", 4)
        ]
        allocations, summary = allocate_to_venues(self.students, venues, seed="SEEDED_123", reshuffle=True)

        self.assertEqual(len(allocations), 2)
        self.assertEqual(len(allocations[0].students), 6)
        self.assertEqual(len(allocations[1].students), 4)
        self.assertEqual(summary["total_students"], 10)
        self.assertEqual(summary["total_capacity"], 10)
        self.assertTrue(summary["is_exact_match"])

        # Verify in-hall students are sorted serially by registration number
        hall1_regs = [s.reg_no for s in allocations[0].students]
        self.assertEqual(hall1_regs, sorted(hall1_regs))

    def test_in_hall_serial_order(self):
        venues = [("HALL A", 5), ("HALL B", 5)]
        allocs, _ = allocate_to_venues(self.students, venues, seed="SEED_XYZ", reshuffle=True, in_hall_order="reg_no")
        for a in allocs:
            regs = [s.reg_no for s in a.students]
            self.assertEqual(regs, sorted(regs))
        sample_text = """
        1\tCOE23HND0120\tKINSMAN ADAVIZE OMANANYI
        2\tCOE23HND0775\tEMMANUEL ORDUE LUMUNGA
        3\tCOE24HND0002\tENOCH TUKURAH ILIYA
        """
        parsed = parse_raw_text(sample_text)
        self.assertEqual(len(parsed), 3)
        self.assertEqual(parsed[0].reg_no, "COE23HND0120")
        self.assertEqual(parsed[0].name, "KINSMAN ADAVIZE OMANANYI")

    def test_parse_existing_docx(self):
        parsed = parse_docx("TUE29_HNDII_CTE442.docx")
        self.assertGreater(len(parsed), 50)
        self.assertEqual(parsed[0].reg_no, "COE23HND0120")

    def test_docx_generation(self):
        venues = [
            ("MINING THEATER-MPRE", 6),
            ("CIVIL ROOM 101/102", 4)
        ]
        allocations, _ = allocate_to_venues(self.students, venues, seed="SEEDED_123")
        docx_bytes = build_attendance_docx(self.config, allocations)
        self.assertGreater(len(docx_bytes), 1000)

        # Verify generated docx can be parsed back
        import docx
        doc = docx.Document(io.BytesIO(docx_bytes))
        self.assertEqual(len(doc.tables), 4)  # 2 metadata tables + 2 student tables

    def test_pdf_generation(self):
        venues = [
            ("MINING THEATER-MPRE", 6),
            ("CIVIL ROOM 101/102", 4)
        ]
        allocations, _ = allocate_to_venues(self.students, venues, seed="SEEDED_123")
        pdf_bytes = build_attendance_pdf(self.config, allocations)
        self.assertGreater(len(pdf_bytes), 1000)
        self.assertTrue(pdf_bytes.startswith(b'%PDF'))


if __name__ == '__main__':
    unittest.main()
