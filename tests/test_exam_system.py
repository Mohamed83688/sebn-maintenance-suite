import unittest
import os
import tempfile
import sqlite3
import json
from docx import Document
from core.exam_manager import ExamManager
from core.user_manager import UserManager
from web_portal.utils.pdf_card_generator import generate_technician_card_pdf

class TestExamSystem(unittest.TestCase):
    def setUp(self):
        # Create temp files for DB and data dirs
        self.db_fd, self.db_path = tempfile.mkstemp()
        self.data_dir = tempfile.mkdtemp()
        
        # Instantiate managers
        self.user_mgr = UserManager(db_path=self.db_path, data_dir=self.data_dir)
        self.exam_mgr = ExamManager(db_path=self.db_path, data_dir=self.data_dir)
        
        # Create test technician
        ok, msg, self.tech_id = self.user_mgr.create_user(
            name="Ahmed Tech",
            username="ahmed.tech",
            password="password123",
            role="TECHNICIAN",
            matricule="TN-TEST",
            shift="A",
            technician_level="LEVEL 25%"
        )
        self.assertTrue(ok)

    def tearDown(self):
        # Clean up temp files
        os.close(self.db_fd)
        try:
            os.unlink(self.db_path)
        except Exception:
            pass
        # remove JSON file from temp data_dir
        json_path = os.path.join(self.data_dir, "Technicians", "TN-TEST.json")
        if os.path.exists(json_path):
            os.remove(json_path)
        try:
            os.rmdir(os.path.join(self.data_dir, "Technicians"))
            os.rmdir(self.data_dir)
        except Exception:
            pass

    def test_exam_creation_and_fields(self):
        parsed_questions = [
            {
                "question_number": 1,
                "question_text": "Quel bouton ouvre le DB Manager ?",
                "choices": [
                    {"answer_text": "DB Manager", "is_correct": 0},
                    {"answer_text": "Log Off", "is_correct": 0},
                    {"answer_text": "Expert Mode", "is_correct": 0}
                ]
            },
            {
                "question_number": 2,
                "question_text": "Quelle est la tension par défaut ?",
                "choices": [
                    {"answer_text": "12V", "is_correct": 0},
                    {"answer_text": "24V", "is_correct": 0}
                ]
            }
        ]

        exam_id = self.exam_mgr.create_exam_with_questions(
            title="CSwin Test 25%",
            description="Examen de base",
            technician_level="LEVEL 25%",
            level_percentage=25,
            source_file="Test25.docx",
            duration=30,
            passing_score=75,
            status="draft",
            parsed_questions=parsed_questions
        )

        self.assertIsNotNone(exam_id)
        
        # Verify exam is draft and correct answers are not defined
        exam = self.exam_mgr.get_exam_by_id(exam_id)
        self.assertEqual(exam['title'], "CSwin Test 25%")
        self.assertEqual(exam['status'], "draft")
        self.assertEqual(len(exam['questions']), 2)
        
        # Verify no correct answer is set (all is_correct are 0)
        for q in exam['questions']:
            corrects = [a for a in q['answers'] if a['is_correct'] == 1]
            self.assertEqual(len(corrects), 0)

    def test_set_correct_answers_and_scoring(self):
        parsed_questions = [
            {
                "question_number": 1,
                "question_text": "Q1",
                "choices": [
                    {"answer_text": "A", "is_correct": 0},
                    {"answer_text": "B", "is_correct": 0}
                ]
            }
        ]

        exam_id = self.exam_mgr.create_exam_with_questions(
            title="CSwin Test 25%",
            description="Examen de base",
            technician_level="LEVEL 25%",
            level_percentage=25,
            source_file="Test25.docx",
            duration=30,
            passing_score=75,
            status="draft",
            parsed_questions=parsed_questions
        )

        exam = self.exam_mgr.get_exam_by_id(exam_id)
        q_id = exam['questions'][0]['id']
        ans_id_correct = exam['questions'][0]['answers'][0]['id']
        ans_id_incorrect = exam['questions'][0]['answers'][1]['id']

        # Configure correct answer
        self.exam_mgr.save_exam_answers_setup(exam_id, {str(q_id): ans_id_correct})

        # Start attempt
        attempt = self.exam_mgr.start_exam_attempt(exam_id, self.tech_id)
        self.assertEqual(attempt['status'], 'running')

        # Save incorrect answer
        self.exam_mgr.save_attempt_answer(attempt['id'], q_id, ans_id_incorrect)
        self.exam_mgr.submit_exam_attempt(attempt['id'])

        # Check attempt result
        with self.exam_mgr._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM exam_attempts WHERE id = ?", (attempt['id'],))
            att = dict(cur.fetchone())
            self.assertEqual(att['status'], 'failed')
            self.assertEqual(att['score'], 0)
            self.assertEqual(att['percentage'], 0.0)

        # Start a new attempt for passed score
        attempt2 = self.exam_mgr.start_exam_attempt(exam_id, self.tech_id)
        self.exam_mgr.save_attempt_answer(attempt2['id'], q_id, ans_id_correct)
        self.exam_mgr.submit_exam_attempt(attempt2['id'])

        with self.exam_mgr._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM exam_attempts WHERE id = ?", (attempt2['id'],))
            att2 = dict(cur.fetchone())
            self.assertEqual(att2['status'], 'passed')
            self.assertEqual(att2['score'], 1)
            self.assertEqual(att2['percentage'], 100.0)

    def test_card_generation(self):
        parsed_questions = [
            {
                "question_number": 1,
                "question_text": "Q1",
                "choices": [
                    {"answer_text": "A", "is_correct": 0}
                ]
            }
        ]
        exam_id = self.exam_mgr.create_exam_with_questions(
            title="CSwin Test 25%",
            description="Examen de base",
            technician_level="LEVEL 25%",
            level_percentage=25,
            source_file="Test25.docx",
            duration=30,
            passing_score=75,
            status="draft",
            parsed_questions=parsed_questions
        )
        exam = self.exam_mgr.get_exam_by_id(exam_id)
        q_id = exam['questions'][0]['id']
        ans_id = exam['questions'][0]['answers'][0]['id']
        self.exam_mgr.save_exam_answers_setup(exam_id, {str(q_id): ans_id})

        # Submit passed attempt
        attempt = self.exam_mgr.start_exam_attempt(exam_id, self.tech_id)
        self.exam_mgr.save_attempt_answer(attempt['id'], q_id, ans_id)
        self.exam_mgr.submit_exam_attempt(attempt['id'])

        # Get card details
        card = self.exam_mgr.get_technician_card_details(self.tech_id)
        self.assertIsNotNone(card)
        self.assertEqual(card['name'], "Ahmed Tech")
        self.assertEqual(card['matricule'], "TN-TEST")
        self.assertEqual(card['status'], "RÉUSSI")
        self.assertEqual(card['percentage'], 100.0)

        # Generate PDF card file
        pdf_path = generate_technician_card_pdf(card)
        self.assertTrue(os.path.exists(pdf_path))
        self.assertGreater(os.path.getsize(pdf_path), 0)
        
        # Clean up PDF file
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

    def test_technician_starts_at_level_0(self):
        # 1. Create a brand new technician without specifying level
        ok, msg, new_tech_id = self.user_mgr.create_user(
            name="Nouveau Tech",
            username="nouveau.tech",
            password="password123",
            role="TECHNICIAN",
            matricule="TN-NEW001",
            shift="B"
        )
        self.assertTrue(ok)
        
        # Verify technician starts at Level 0
        tech_user = self.user_mgr.get_user_by_id(new_tech_id)
        self.assertEqual(tech_user['technician_level'], 'Level 0')

        # Check progression object
        prog = self.exam_mgr.get_technician_progression(new_tech_id)
        self.assertEqual(prog['current_level'], 'Level 0')
        self.assertEqual(prog['next_level'], 'Level 1')

        # 2. Create and pass a Level 1 exam
        parsed_questions = [{
            "question_number": 1,
            "question_text": "Q1 Level 1",
            "choices": [{"answer_text": "Correct Ans", "is_correct": 1}]
        }]
        exam_id = self.exam_mgr.create_exam_with_questions(
            title="Examen Level 1 Test",
            description="Examen Level 1",
            technician_level="Level 1",
            level_percentage=25,
            source_file="TestL1.docx",
            duration=30,
            passing_score=75,
            status="published",
            parsed_questions=parsed_questions
        )
        exam = self.exam_mgr.get_exam_by_id(exam_id)
        q_id = exam['questions'][0]['id']
        ans_id = exam['questions'][0]['answers'][0]['id']
        self.exam_mgr.save_exam_answers_setup(exam_id, {str(q_id): ans_id})

        # Start attempt and pass
        attempt = self.exam_mgr.start_exam_attempt(exam_id, new_tech_id)
        self.exam_mgr.save_attempt_answer(attempt['id'], q_id, ans_id)
        status = self.exam_mgr.submit_exam_attempt(attempt['id'])
        self.assertEqual(status, 'passed')

        # 3. Technician should now be advanced from Level 0 to Level 1
        tech_user_after = self.user_mgr.get_user_by_id(new_tech_id)
        self.assertEqual(tech_user_after['technician_level'], 'Level 1')

if __name__ == '__main__':
    unittest.main()
