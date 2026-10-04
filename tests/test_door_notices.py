import unittest
from core.models import Student, ExamConfig, VenueAllocation
from core.door_notices import build_door_notices_docx, build_door_notices_pdf


class TestDoorNotices(unittest.TestCase):
    def setUp(self):
        self.students = [
            Student(reg_no=f"COE25HND{i:04d}", name=f"STUDENT NAME {i}")
            for i in range(1, 41)
        ]
        self.config = ExamConfig(
            course_title="ARTIFICIAL INTELLIGENCE",
            course_code="CTE 442",
            exam_date="29th Sept, 2026",
            level="HND II"
        )
        self.allocations = [
            VenueAllocation(venue_name="MINING THEATER-MPRE", capacity=25, students=self.students[:25]),
            VenueAllocation(venue_name="CIVIL ROOM 101/102", capacity=15, students=self.students[25:])
        ]

    def test_door_notices_docx(self):
        data = build_door_notices_docx(self.config, self.allocations)
        self.assertGreater(len(data), 1000)

    def test_door_notices_pdf(self):
        data = build_door_notices_pdf(self.config, self.allocations)
        self.assertGreater(len(data), 1000)
        self.assertTrue(data.startswith(b'%PDF'))


if __name__ == '__main__':
    unittest.main()
