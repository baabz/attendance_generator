import unittest
from core.models import Student, ExamConfig, VenueAllocation
from core.envelope_tags import build_envelope_tags_docx, build_envelope_tags_pdf


class TestEnvelopeTags(unittest.TestCase):
    def setUp(self):
        self.students = [
            Student(reg_no=f"COE25HND{i:04d}", name=f"STUDENT NAME {i}")
            for i in range(1, 31)
        ]
        self.config = ExamConfig(
            course_title="ARTIFICIAL INTELLIGENCE",
            course_code="CTE 442",
            exam_date="29th Sept, 2026",
            level="HND II"
        )
        self.allocations = [
            VenueAllocation(venue_name="MINING THEATER-MPRE", capacity=20, students=self.students[:20]),
            VenueAllocation(venue_name="CIVIL ROOM 101/102", capacity=10, students=self.students[20:])
        ]

    def test_envelope_tags_docx(self):
        data = build_envelope_tags_docx(self.config, self.allocations)
        self.assertGreater(len(data), 1000)

    def test_envelope_tags_pdf(self):
        data = build_envelope_tags_pdf(self.config, self.allocations)
        self.assertGreater(len(data), 1000)
        self.assertTrue(data.startswith(b'%PDF'))


if __name__ == '__main__':
    unittest.main()
