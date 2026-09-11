import os
import sqlite3
import datetime
import re
import json
import uuid
import logging
try:
    from docx import Document
except ImportError:
    Document = None

logger = logging.getLogger('sebn-maintenance')

DEFAULT_CSWIN_Q6_BOXES = [
    {"id": 1, "num": 1, "target": "Start Test Mode", "left": 5.13, "top": 19.33, "width": 14.45, "height": 10.81, "hint": "Test Mode (▶)"},
    {"id": 2, "num": 2, "target": "Log Off registered User", "left": 23.38, "top": 1.66, "width": 20.44, "height": 12.06, "hint": "Log Off (🔓)"},
    {"id": 3, "num": 3, "target": "Open Setup CS WIN nx", "left": 47.24, "top": 0.21, "width": 17.68, "height": 13.72, "hint": "Setup (🔧)"},
    {"id": 4, "num": 4, "target": "Open Expert Mode", "left": 68.44, "top": 0.42, "width": 19.39, "height": 13.51, "hint": "Expert Mode (👥)"},
    {"id": 5, "num": 5, "target": "Start CS WIN nx Helpsheet", "left": 81.75, "top": 17.88, "width": 16.35, "height": 11.43, "hint": "Helpsheet (?)"},
    {"id": 6, "num": 6, "target": "Edit / Create TesterConfiguration", "left": 1.90, "top": 53.85, "width": 17.40, "height": 12.06, "hint": "Configure Test Environment"},
    {"id": 7, "num": 7, "target": "Edit / Create Programs", "left": 13.02, "top": 88.98, "width": 19.87, "height": 8.32, "hint": "Edit My Programs"},
    {"id": 8, "num": 8, "target": "Edit / Create Connector Library", "left": 34.98, "top": 88.36, "width": 19.30, "height": 10.81, "hint": "Edit My Components"},
    {"id": 9, "num": 9, "target": "Edit / Create Labels", "left": 58.17, "top": 89.19, "width": 16.63, "height": 10.60, "hint": "Edit My Labels"},
    {"id": 10, "num": 10, "target": "Edit / Create Projects", "left": 81.75, "top": 57.80, "width": 16.54, "height": 11.23, "hint": "Edit My Projects"},
    {"id": 11, "num": 11, "target": "Close CS WIN nx", "left": 81.37, "top": 77.96, "width": 17.21, "height": 9.36, "hint": "Quitter (✖)"}
]

DEFAULT_CSWIN_Q6_LABELS = [
    "Start Test Mode",
    "Log Off registered User",
    "Open Setup CS WIN nx",
    "Open Expert Mode",
    "Start CS WIN nx Helpsheet",
    "Edit / Create TesterConfiguration",
    "Edit / Create Programs",
    "Edit / Create Connector Library",
    "Edit / Create Labels",
    "Edit / Create Projects",
    "Close CS WIN nx"
]


class ExamManager:
    """
    Manages Exams, QCM Questions, Options, Attempts, and Scoring for SEBN-TN.
    Integrates directly with the centralized SQLite database (IMA.db).
    """
    def __init__(self, db_path: str, data_dir: str):
        self.db_path = db_path
        self.data_dir = data_dir
        self._init_db()
        self.auto_import_all_docx()

    def _get_conn(self):
        conn = sqlite3.connect(self.db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        os.makedirs(os.path.dirname(os.path.abspath(self.db_path)), exist_ok=True)
        with self._get_conn() as conn:
            cur = conn.cursor()
            
            # 1. Exams Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS exams (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    title TEXT NOT NULL,
                    description TEXT,
                    technician_level TEXT NOT NULL,
                    level_percentage INTEGER NOT NULL,
                    source_file TEXT,
                    duration INTEGER NOT NULL DEFAULT 30,
                    number_of_questions INTEGER DEFAULT 0,
                    passing_score INTEGER DEFAULT 75, -- target percentage required to pass
                    status TEXT DEFAULT 'draft',      -- 'draft', 'published', 'archived'
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # 2. Questions Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS exam_questions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exam_id INTEGER NOT NULL,
                    question_number INTEGER NOT NULL,
                    question_text TEXT NOT NULL,
                    question_type TEXT DEFAULT 'multiple_choice',
                    image TEXT,
                    available_labels TEXT,  -- JSON array for image_labeling type
                    FOREIGN KEY(exam_id) REFERENCES exams(id) ON DELETE CASCADE
                )
            """)

            # 3. Answers Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS exam_answers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    question_id INTEGER NOT NULL,
                    answer_text TEXT NOT NULL,
                    is_correct INTEGER DEFAULT 0, -- 0 or 1
                    FOREIGN KEY(question_id) REFERENCES exam_questions(id) ON DELETE CASCADE
                )
            """)

            # 4. Attempts Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS exam_attempts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exam_id INTEGER NOT NULL,
                    user_id INTEGER NOT NULL,
                    started_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    submitted_at DATETIME,
                    score INTEGER,
                    percentage REAL,
                    correct_answers INTEGER,
                    incorrect_answers INTEGER,
                    unanswered INTEGER,
                    status TEXT DEFAULT 'running', -- 'running', 'passed', 'failed'
                    FOREIGN KEY(exam_id) REFERENCES exams(id) ON DELETE CASCADE,
                    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
                )
            """)

            # 5. Attempt Answers Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS exam_attempt_answers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    attempt_id INTEGER NOT NULL,
                    question_id INTEGER NOT NULL,
                    selected_answer_id INTEGER,
                    FOREIGN KEY(attempt_id) REFERENCES exam_attempts(id) ON DELETE CASCADE,
                    FOREIGN KEY(question_id) REFERENCES exam_questions(id) ON DELETE CASCADE
                )
            """)

            # 6. Level Change Audit Log
            cur.execute("""
                CREATE TABLE IF NOT EXISTS level_change_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    previous_level TEXT,
                    new_level TEXT NOT NULL,
                    changed_by TEXT NOT NULL,
                    change_type TEXT DEFAULT 'EXAM',
                    reason TEXT,
                    exam_attempt_id INTEGER,
                    changed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
                )
            """)

            # 7. Extra question images (supports multiple images per question)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS exam_question_images (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    question_id INTEGER NOT NULL,
                    filename TEXT NOT NULL,
                    sort_order INTEGER DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(question_id) REFERENCES exam_questions(id) ON DELETE CASCADE
                )
            """)

            conn.commit()
            self._migrate_db(conn)

    def _migrate_db(self, conn):
        """Applies incremental schema migrations for existing databases."""
        cur = conn.cursor()

        # Migration 1: exam_question_images table
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='exam_question_images'")
        if not cur.fetchone():
            cur.execute("""
                CREATE TABLE exam_question_images (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    question_id INTEGER NOT NULL,
                    filename TEXT NOT NULL,
                    sort_order INTEGER DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(question_id) REFERENCES exam_questions(id) ON DELETE CASCADE
                )
            """)
            conn.commit()
            logger.info("[DB] Migration: created exam_question_images table.")

        # Migration 2: question_type column (rename legacy 'qcm' → 'multiple_choice')
        cur.execute("PRAGMA table_info(exam_questions)")
        cols = {r['name'] for r in cur.fetchall()}
        if 'question_type' not in cols:
            cur.execute("ALTER TABLE exam_questions ADD COLUMN question_type TEXT DEFAULT 'multiple_choice'")
            conn.commit()
            logger.info("[DB] Migration: added question_type column.")
        else:
            # Normalise legacy 'qcm' value
            cur.execute("UPDATE exam_questions SET question_type='multiple_choice' WHERE question_type='qcm' OR question_type IS NULL")
            conn.commit()

        # Migration 3: available_labels column
        cur.execute("PRAGMA table_info(exam_questions)")
        cols = {r['name'] for r in cur.fetchall()}
        if 'available_labels' not in cols:
            cur.execute("ALTER TABLE exam_questions ADD COLUMN available_labels TEXT")
            conn.commit()
            logger.info("[DB] Migration: added available_labels column.")

        # Migration 4: boxes_json column on exam_questions
        cur.execute("PRAGMA table_info(exam_questions)")
        cols = {r['name'] for r in cur.fetchall()}
        if 'boxes_json' not in cols:
            cur.execute("ALTER TABLE exam_questions ADD COLUMN boxes_json TEXT")
            conn.commit()
            logger.info("[DB] Migration: added boxes_json column.")

        # Migration 5: answer_text column on exam_attempt_answers
        cur.execute("PRAGMA table_info(exam_attempt_answers)")
        cols_ans = {r['name'] for r in cur.fetchall()}
        if 'answer_text' not in cols_ans:
            cur.execute("ALTER TABLE exam_attempt_answers ADD COLUMN answer_text TEXT")
            conn.commit()
            logger.info("[DB] Migration: added answer_text column.")

        # Migration 6: Repair and ensure all image_labeling / CS WIN Q6 questions have boxes_json and available_labels set
        try:
            cur.execute("""
                SELECT id, question_number, question_text, question_type, available_labels, boxes_json 
                FROM exam_questions 
                WHERE question_type = 'image_labeling' 
                   OR question_text LIKE '%cadre vide%' 
                   OR question_text LIKE '%remplissez le cadre%'
            """)
            q_rows = cur.fetchall()
            for qr in q_rows:
                qid = qr['id']
                raw_b = qr['boxes_json']
                raw_l = qr['available_labels']
                q_text = (qr['question_text'] or '').lower()
                
                if 'cadre' in q_text:
                    needs_update = False
                    new_b = raw_b
                    new_l = raw_l
                    if not raw_b or len(str(raw_b)) < 20:
                        new_b = json.dumps(DEFAULT_CSWIN_Q6_BOXES, ensure_ascii=False)
                        needs_update = True
                    has_merged = any(k in str(raw_l) for k in ['Log Off registered User/Start', 'Create Projects/Close', 'TesterConfiguration/Open', 'Libary/Edit', 'Labels/Edit'])
                    if not raw_l or len(str(raw_l)) < 20 or has_merged or '/' in str(raw_l):
                        new_l = json.dumps(DEFAULT_CSWIN_Q6_LABELS, ensure_ascii=False)
                        needs_update = True
                    if needs_update or qr['question_type'] != 'image_labeling':
                        cur.execute("""
                            UPDATE exam_questions 
                            SET question_type = 'image_labeling', available_labels = ?, boxes_json = ? 
                            WHERE id = ?
                        """, (new_l, new_b, qid))

            # Ensure correct answers are set for CS WIN QCM questions if none are marked
            correct_keywords = {
                1: 'log off',
                2: 'tsk',
                3: 'porta',
                4: 'db manager',
                5: 'user manager',
                7: 'network range',
                8: 'test point cards',
                9: 'pin definitions',
                10: 'ok or not ok',
                11: 'closed for ok',
                12: 'current status of the switches',
                13: 'enter a new name',
                14: 'sks',
                15: '24v',
                16: 'force de ressort',
                17: 'seiri'
            }
            cur.execute("""
                SELECT DISTINCT exam_id FROM exam_questions WHERE question_number = 6 AND (question_type = 'image_labeling' OR question_text LIKE '%cadre%')
            """)
            cswin_exam_ids = [r['exam_id'] for r in cur.fetchall()]
            for eid in cswin_exam_ids:
                cur.execute("""
                    SELECT SUM(a.is_correct) FROM exam_questions q 
                    JOIN exam_answers a ON a.question_id = q.id 
                    WHERE q.exam_id = ?
                """, (eid,))
                sum_row = cur.fetchone()
                if not sum_row or not sum_row[0] or sum_row[0] == 0:
                    for qnum, kw in correct_keywords.items():
                        cur.execute("""
                            SELECT a.id, a.answer_text FROM exam_questions q 
                            JOIN exam_answers a ON a.question_id = q.id 
                            WHERE q.exam_id = ? AND q.question_number = ?
                        """, (eid, qnum))
                        ans_list = cur.fetchall()
                        for aid, atxt in ans_list:
                            if kw in (atxt or '').lower():
                                cur.execute("UPDATE exam_answers SET is_correct = 1 WHERE id = ?", (aid,))
                                break
            conn.commit()
            logger.info("[DB] Migration 6: checked and repaired image_labeling questions and answers.")
        except Exception as e:
            logger.warning(f"[DB] Migration 6 warning: {e}")


    # ── Parsing & Auto-import Logic ──────────────────────────────────────────

    # ── Keywords that signal image-labeling (not QCM) ──────────────────────
    _LABELING_KEYWORDS = [
        'éléments disponibles', 'elements disponibles',
        'remplissez le cadre', 'remplissez les cases', 'remplissez les cadres',
        'complétez les cases', 'completez les cases',
        'placez les réponses', 'placez les etiquettes', 'placez les étiquettes',
        'glissez les', 'déposez les', 'deposez les',
        'word bank', 'bank of words', 'answer bank',
        'associez les', 'reliez les',
    ]

    @classmethod
    def _detect_question_type(cls, question_text: str, choices: list, has_image: bool,
                               has_txbx: bool) -> str:
        """
        Automatically determine question type from parser signals.
        Returns: 'multiple_choice' | 'image_labeling' | 'true_false' | 'text_input'
        """
        text_low = question_text.lower()

        # 1. Image labeling: explicit keyword + has an image or textbox labels
        has_labeling_kw = any(kw in text_low for kw in cls._LABELING_KEYWORDS)
        if has_labeling_kw and (has_image or has_txbx):
            return 'image_labeling'

        # 2. True / False: exactly 2 choices that are vrai/faux variants
        if len(choices) == 2:
            texts = {c['answer_text'].strip().lower() for c in choices}
            tf_pairs = [
                {'vrai', 'faux'}, {'true', 'false'},
                {'oui', 'non'}, {'yes', 'no'},
                {'correct', 'incorrect'}, {'juste', 'faux'},
            ]
            if any(texts == pair for pair in tf_pairs):
                return 'true_false'

        # 3. Standard QCM: A/B/C/D choices present
        if choices:
            return 'multiple_choice'

        # 4. Fallback: no choices, no image → treat as observational / text_input
        return 'text_input'

    def parse_docx_exam(self, file_path: str, output_images_dir: str = None):
        """
        Parses questions, multiple choice options, and images from a docx file.
        Document-structure-aware two-pass parser:
          - Pass 1: Extract paragraph data (text, images, textboxes, Q-header detection)
          - Pass 2: Assign images to questions by boundary, classify question types

        Returns a list of question dicts:
          [{
             'question_number': int,
             'question_text': str,
             'question_type': str,             # 'multiple_choice' | 'image_labeling' | ...
             'choices': [{'answer_text': str, 'is_correct': int}],  # QCM / true_false only
             'available_labels': [str],         # image_labeling only
             'image': str or None,
             'extra_images': [str]
          }]
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        if Document is None:
            raise ImportError("Le module python-docx est requis pour analyser les fichiers Word (.docx). Installez-le avec: pip install python-docx")

        doc = Document(file_path)

        def normalize_text(t):
            if not t:
                return ""
            return re.sub(r'[\u2010\u2011\u2012\u2013\u2014\u2212]', '-', t)

        if output_images_dir:
            os.makedirs(output_images_dir, exist_ok=True)

        # ── PASS 1: Collect paragraph data ────────────────────────────────
        para_data = []

        for p_idx, p in enumerate(doc.paragraphs):
            p_text = normalize_text(p.text.strip())

            # Extract textbox paragraphs (answer banks, label lists)
            txbx_paras = p._element.xpath('.//*[local-name()="txbxContent"]//*[local-name()="p"]')
            seen_t = []
            for tp in txbx_paras:
                texts = [t.text for t in tp.xpath('.//*[local-name()="t"]') if t.text]
                full = normalize_text(''.join(texts).strip())
                if full and full not in seen_t:
                    seen_t.append(full)
            txbx_text = '\n'.join(seen_t)

            # Extract blips (images)
            p_images = []
            for c_idx, child in enumerate(p._element):
                blips = child.xpath('.//*[local-name()="blip"]')
                for b in blips:
                    for k, v in b.attrib.items():
                        if 'embed' in k:
                            rel = doc.part.rels.get(v)
                            if rel and hasattr(rel, 'target_part') and rel.target_part:
                                filename = os.path.basename(rel.target_ref)
                                img_bytes = rel.target_part.blob
                                saved_filename = filename
                                if output_images_dir:
                                    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'png'
                                    saved_filename = f"q_{uuid.uuid4().hex[:10]}.{ext}"
                                    out_path = os.path.join(output_images_dir, saved_filename)
                                    with open(out_path, 'wb') as img_out:
                                        img_out.write(img_bytes)
                                p_images.append({'c_idx': c_idx, 'filename': saved_filename})

            # Detect question header (Q1, Q2, ... style)
            run_q_match = None
            run_header_idx = -1
            for r_i, child in enumerate(p._element):
                ctext = normalize_text(''.join(child.itertext()).strip())
                rm = re.search(r'(?:^|\b)Q\s*(\d+)', ctext, re.IGNORECASE)
                if rm:
                    run_q_match = rm
                    run_header_idx = r_i
                    break

            q_match = re.search(r'(?:^|\b)Q\s*(\d+)[\.\:\-\s]+(.*)', p_text, re.DOTALL)
            q_num = None
            header_text = ''
            if q_match:
                q_num = int(q_match.group(1))
                header_text = q_match.group(2).strip()
            elif run_q_match:
                q_num = int(run_q_match.group(1))
                header_text = p_text

            para_data.append({
                'idx': p_idx,
                'text': p_text,
                'txbx_text': txbx_text,
                'images': p_images,
                'q_num': q_num,
                'header_text': header_text,
                'run_header_idx': run_header_idx
            })

        # ── PASS 2: Build questions, assign images, collect choices ───────
        questions = []
        current_q = None

        def _add_choice_line(q, line):
            opt_match = re.match(r'^(?:[A-Da-d][\.\)\:\-]|[1-4][\.\)])\s*(.*)', line)
            if opt_match:
                q['choices'].append({'answer_text': opt_match.group(1).strip(), 'is_correct': 0})
                return True
            return False

        q_seq = 0
        for p in para_data:
            # Check for non-exam section boundaries
            p_text_low = p['text'].lower()
            if p_text_low.startswith(('exercice pratique', 'critères d\'évaluation', 'criteres d\'evaluation')):
                current_q = None
                continue

            if p['q_num'] is not None:
                q_seq += 1
                raw_header = p['header_text'] or p['text']
                # If header has embedded newlines with choices (e.g. Q9, Q14-Q16)
                header_lines = [l.strip() for l in raw_header.split('\n') if l.strip()]
                q_text_lines = []
                choices = []

                for hl in header_lines:
                    opt_m = re.match(r'^(?:[A-Da-d][\.\)\:\-]|[1-4][\.\)])\s*(.*)', hl)
                    if opt_m:
                        choices.append({'answer_text': opt_m.group(1).strip(), 'is_correct': 0})
                    elif not choices:
                        q_text_lines.append(hl)
                    else:
                        choices[-1]['answer_text'] += ' ' + hl

                current_q = {
                    'question_number': q_seq,
                    'question_text': '\n'.join(q_text_lines) if q_text_lines else raw_header,
                    'choices': choices,
                    'available_labels': [],
                    'image': None,
                    'extra_images': [],
                    '_all_images': [],
                    '_txbx_lines': list(filter(None, p['txbx_text'].split('\n'))) if p['txbx_text'] else []
                }
                questions.append(current_q)

                # Image-boundary logic: image that appears in the same paragraph
                # as the question header but BEFORE the header run → belongs to prev question
                if p['images']:
                    prev_q = questions[-2] if len(questions) >= 2 else None
                    for img in p['images']:
                        assigned_to_prev = False
                        if img['c_idx'] < p['run_header_idx'] and prev_q and not prev_q['_all_images']:
                            low_prev = prev_q['question_text'].lower()
                            if any(kw in low_prev for kw in [
                                'image ci-dessous', 'image ci dessous', 'figure ci-dessous',
                                'examinez l\'image', 'examinez', 'voir ci-dessous'
                            ]):
                                prev_q['_all_images'].append(img['filename'])
                                assigned_to_prev = True
                        if not assigned_to_prev:
                            current_q['_all_images'].append(img['filename'])
            else:
                if current_q is None:
                    continue

                # Collect images
                for img in p['images']:
                    current_q['_all_images'].append(img['filename'])

                # Collect textbox labels
                if p['txbx_text']:
                    new_lines = [l.strip() for l in p['txbx_text'].split('\n') if l.strip()]
                    current_q['_txbx_lines'].extend(new_lines)

                # Parse lines for choices or body text
                if p['text']:
                    lines = [l.strip() for l in p['text'].split('\n') if l.strip()]
                    for line in lines:
                        added = _add_choice_line(current_q, line)
                        if not added:
                            if not current_q['choices']:
                                # If question ends with ? or :, treat following lines as un-prefixed choices
                                if current_q['question_text'].strip().endswith(('?', ':')):
                                    current_q['choices'].append({'answer_text': line, 'is_correct': 0})
                                else:
                                    current_q['question_text'] += '\n' + line
                            elif len(current_q['choices']) < 5 and current_q['question_text'].strip().endswith(('?', ':')):
                                current_q['choices'].append({'answer_text': line, 'is_correct': 0})
                            else:
                                current_q['choices'][-1]['answer_text'] += ' ' + line


        # ── POST-PROCESSING: types, images, label extraction ──────────────
        for q in questions:
            # Resolve images
            if q['_all_images']:
                q['image'] = q['_all_images'][0]
                q['extra_images'] = q['_all_images'][1:]
            else:
                q['image'] = None
                q['extra_images'] = []
            del q['_all_images']

            has_image = bool(q['image'])
            has_txbx = bool(q['_txbx_lines'])

            # Classify question type
            q['question_type'] = self._detect_question_type(
                q['question_text'], q['choices'], has_image, has_txbx
            )

            # For image_labeling: convert choices → available_labels
            # and pull in any textbox lines as additional labels
            if q['question_type'] == 'image_labeling':
                # For CS WIN Question 6 (or image labeling with cadre vide), use 11 clean labels and predefined box coordinates
                if q.get('question_number') == 6 or 'cadre vide' in q.get('question_text', '').lower():
                    q['available_labels'] = list(DEFAULT_CSWIN_Q6_LABELS)
                    q['boxes_json'] = json.dumps(DEFAULT_CSWIN_Q6_BOXES, ensure_ascii=False)
                else:
                    # Gather label candidates: A/B/C/D lines + textbox lines
                    label_set = []
                    for c in q['choices']:
                        label_set.append(c['answer_text'])
                    for tl in q['_txbx_lines']:
                        # Textbox lines may be slash-separated: split them
                        for part in tl.split(' / '):
                            part = part.strip().rstrip('/')
                            if part and part not in label_set:
                                label_set.append(part)
                    q['available_labels'] = label_set
                q['choices'] = []   # NOT radio choices
            else:
                q['available_labels'] = []
                # Append textbox content to question body for non-labeling types
                txbx_joined = '\n'.join(q['_txbx_lines'])
                if txbx_joined and txbx_joined not in q['question_text']:
                    q['question_text'] += f"\n\n[Éléments disponibles :]\n{txbx_joined}"

            del q['_txbx_lines']

        return questions

    def auto_import_all_docx(self):
        """
        Detects the 4 real test docx files in downloads and auto-imports them to database
        if they are not already imported.
        """
        download_dir = r"C:\Users\hama0\Downloads"
        if not os.path.isdir(download_dir):
            return

        levels_map = {
            "25%": {"title": "CSwin Basic Knowledge", "pct": 25, "level": "Level 1"},
            "50%": {"title": "CSwin Creation Hardware", "pct": 50, "level": "Level 2"},
            "75%": {"title": "CSwin & Brainware & Vacuum", "pct": 75, "level": "Level 3"},
            "100%": {"title": "CSwin & Brainware & Vacuum", "pct": 100, "level": "Level 4"}
        }

        # Resolve output image dir
        img_out = os.path.join(self.data_dir, "exam_images")
        if not os.path.exists(img_out):
            # check project/web_portal/static/exam_images
            alt_img = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "web_portal", "static", "exam_images"))
            if os.path.exists(alt_img):
                img_out = alt_img

        try:
            files = os.listdir(download_dir)
        except Exception:
            return

        for f in files:
            if "Test Level" in f and f.endswith(".docx"):
                lvl_key = None
                for key in levels_map:
                    if key in f:
                        lvl_key = key
                        break
                
                if lvl_key:
                    cfg = levels_map[lvl_key]
                    full_path = os.path.join(download_dir, f)
                    
                    # Check if already imported
                    if self.is_exam_imported(f):
                        continue
                        
                    try:
                        logger.info(f"Auto-importing exam {f} for {cfg['level']}")
                        questions = self.parse_docx_exam(full_path, output_images_dir=img_out)
                        self.create_exam_with_questions(
                            title=cfg["title"],
                            description=f"Examen officiel pour le niveau {cfg['level']}",
                            technician_level=cfg["level"],
                            level_percentage=cfg["pct"],
                            source_file=f,
                            duration=30,
                            passing_score=75,
                            status='draft',
                            parsed_questions=questions
                        )
                    except Exception as e:
                        logger.error(f"Error auto-importing exam file {f}: {e}")

    def is_exam_imported(self, source_file: str) -> bool:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id FROM exams WHERE source_file = ?", (source_file,))
            return cur.fetchone() is not None

    def create_exam_with_questions(self, title, description, technician_level, level_percentage, source_file, duration, passing_score, status, parsed_questions):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO exams (title, description, technician_level, level_percentage, source_file, duration, number_of_questions, passing_score, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (title, description, technician_level, level_percentage, source_file, duration, len(parsed_questions), passing_score, status))
            exam_id = cur.lastrowid

            import json as _json
            for q in parsed_questions:
                q_type = q.get('question_type', 'multiple_choice')
                labels = q.get('available_labels', [])
                if q_type == 'image_labeling' and (not labels or len(labels) < 10):
                    labels = DEFAULT_CSWIN_Q6_LABELS
                labels_json = _json.dumps(labels, ensure_ascii=False) if labels else None

                boxes = q.get('boxes_json')
                if q_type == 'image_labeling' and not boxes:
                    boxes = _json.dumps(DEFAULT_CSWIN_Q6_BOXES, ensure_ascii=False)

                cur.execute("""
                    INSERT INTO exam_questions (exam_id, question_number, question_text, question_type, image, available_labels, boxes_json)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (exam_id, q['question_number'], q['question_text'], q_type, q.get('image'), labels_json, boxes))
                q_id = cur.lastrowid

                # Persist extra images
                for sort_i, ef in enumerate(q.get('extra_images', [])):
                    cur.execute("""
                        INSERT INTO exam_question_images (question_id, filename, sort_order)
                        VALUES (?, ?, ?)
                    """, (q_id, ef, sort_i + 1))

                # Only insert exam_answers for gradable types
                if q_type in ('multiple_choice', 'true_false', 'qcm'):
                    for choice in q.get('choices', []):
                        cur.execute("""
                            INSERT INTO exam_answers (question_id, answer_text, is_correct)
                            VALUES (?, ?, ?)
                        """, (q_id, choice['answer_text'], choice.get('is_correct', 0)))
            conn.commit()
            return exam_id

    # ── Exam Retrieval and Management ────────────────────────────────────────

    def get_all_exams(self, role=None):
        with self._get_conn() as conn:
            cur = conn.cursor()
            if role in ('OWNER', 'ADMIN', 'admin', 'OWNER'):
                cur.execute("SELECT * FROM exams ORDER BY level_percentage ASC, title ASC")
            else:
                cur.execute("SELECT * FROM exams WHERE status = 'published' ORDER BY level_percentage ASC, title ASC")
            exams = [dict(r) for r in cur.fetchall()]
            # Enrich with correct_count so admin UI can warn when no answers are configured
            for ex in exams:
                cur.execute("""
                    SELECT COUNT(*) as cnt
                    FROM exam_answers ea
                    JOIN exam_questions eq ON eq.id = ea.question_id
                    WHERE eq.exam_id = ? AND ea.is_correct = 1
                """, (ex['id'],))
                row = cur.fetchone()
                ex['correct_count'] = row['cnt'] if row else 0
            return exams

    def get_exam_by_id(self, exam_id: int):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM exams WHERE id = ?", (exam_id,))
            row = cur.fetchone()
            if not row:
                return None
            
            exam = dict(row)
            # Fetch questions
            cur.execute("SELECT * FROM exam_questions WHERE exam_id = ? ORDER BY question_number ASC", (exam_id,))
            qs = [dict(q) for q in cur.fetchall()]
            
            import json as _json
            for q in qs:
                cur.execute("SELECT * FROM exam_answers WHERE question_id = ?", (q['id'],))
                q['answers'] = [dict(a) for a in cur.fetchall()]

                # Decode available_labels JSON
                raw_labels = q.get('available_labels')
                if raw_labels:
                    try:
                        q['available_labels'] = _json.loads(raw_labels)
                    except Exception:
                        q['available_labels'] = []
                else:
                    q['available_labels'] = []

                # Decode boxes_json for image labeling
                raw_boxes = q.get('boxes_json')
                if raw_boxes:
                    try:
                        q['labeling_boxes'] = _json.loads(raw_boxes)
                    except Exception:
                        q['labeling_boxes'] = []
                else:
                    q['labeling_boxes'] = []

                # Check if this question is CS WIN Question 6 (or image labeling with cadre vide)
                q_text_low = (q.get('question_text') or '').lower()
                is_cswin_q6 = (
                    q.get('question_type') == 'image_labeling'
                    or 'cadre vide' in q_text_low
                    or (q.get('question_number') == 6 and any(k in q_text_low for k in ['cadre', 'réponse', 'reponse', 'ci-dessous']))
                    or (q.get('question_number') == 6 and 'cswin' in (exam.get('title') or '').lower())
                )

                if is_cswin_q6:
                    q['question_type'] = 'image_labeling'
                    if not q['labeling_boxes']:
                        q['labeling_boxes'] = list(DEFAULT_CSWIN_Q6_BOXES)
                    # Clean available_labels if empty or contains merged labels (e.g. slashes/Edit combined)
                    has_merged_labels = any('/' in str(lbl) and 'Edit' in str(lbl) for lbl in q['available_labels'])
                    if not q['available_labels'] or len(q['available_labels']) < 10 or has_merged_labels:
                        q['available_labels'] = list(DEFAULT_CSWIN_Q6_LABELS)

                # Normalise legacy question_type
                if not q.get('question_type') or q['question_type'] == 'qcm':
                    q['question_type'] = 'multiple_choice'

                # Build combined images list
                extra = self._get_question_extra_images_cur(cur, q['id'])
                imgs = []
                if q.get('image'):
                    imgs.append(q['image'])
                for ei in extra:
                    if ei['filename'] not in imgs:
                        imgs.append(ei['filename'])
                q['images'] = imgs

            exam['questions'] = qs
            return exam

    def update_exam_details(self, exam_id: int, title: str, description: str, duration: int, passing_score: int, status: str):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE exams 
                SET title = ?, description = ?, duration = ?, passing_score = ?, status = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (title, description, duration, passing_score, status, exam_id))
            conn.commit()
            return True

    def save_exam_answers_setup(self, exam_id: int, correct_answers_map: dict):
        """
        Updates the correct choices in database based on correct_answers_map.
        correct_answers_map is a dict of question_id -> correct_answer_id.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            
            # Fetch all questions for this exam
            cur.execute("SELECT id FROM exam_questions WHERE exam_id = ?", (exam_id,))
            q_ids = [r['id'] for r in cur.fetchall()]
            
            for q_id in q_ids:
                # Set all answers to incorrect first
                cur.execute("UPDATE exam_answers SET is_correct = 0 WHERE question_id = ?", (q_id,))
                
                # Update the selected correct answer
                correct_ans_id = correct_answers_map.get(str(q_id)) or correct_answers_map.get(q_id)
                if correct_ans_id:
                    cur.execute("UPDATE exam_answers SET is_correct = 1 WHERE question_id = ? AND id = ?", (q_id, correct_ans_id))
            
            conn.commit()
            return True

    def delete_exam(self, exam_id: int):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM exams WHERE id = ?", (exam_id,))
            conn.commit()
            return True

    # ── Exam Attempts & Taking Logic ─────────────────────────────────────────

    def start_exam_attempt(self, exam_id: int, user_id: int):
        """
        Starts an attempt. If an active attempt exists, returns it (to persist timer on reload).
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            
            # Check for existing running attempt
            cur.execute("""
                SELECT * FROM exam_attempts 
                WHERE exam_id = ? AND user_id = ? AND status = 'running'
            """, (exam_id, user_id))
            existing = cur.fetchone()
            if existing:
                return dict(existing)
            
            # Start a new one
            cur.execute("""
                INSERT INTO exam_attempts (exam_id, user_id, started_at, status)
                VALUES (?, ?, datetime('now', 'localtime'), 'running')
            """, (exam_id, user_id))
            conn.commit()
            
            cur.execute("SELECT * FROM exam_attempts WHERE id = ?", (cur.lastrowid,))
            return dict(cur.fetchone())

    def get_active_attempt(self, attempt_id: int):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM exam_attempts WHERE id = ?", (attempt_id,))
            row = cur.fetchone()
            return dict(row) if row else None

    def get_remaining_seconds(self, attempt_id: int, duration_minutes: int) -> int:
        attempt = self.get_active_attempt(attempt_id)
        if not attempt or attempt['status'] != 'running':
            return 0
            
        started_str = attempt['started_at']
        # Started date is saved in sqlite format YYYY-MM-DD HH:MM:SS
        try:
            started_dt = datetime.datetime.strptime(started_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            started_dt = datetime.datetime.now()
            
        elapsed = (datetime.datetime.now() - started_dt).total_seconds()
        total_allowed = duration_minutes * 60
        rem = int(total_allowed - elapsed)
        return max(0, rem)

    def save_attempt_answer(self, attempt_id: int, question_id: int, selected_answer_id: int = None, answer_text: str = None):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id FROM exam_attempt_answers 
                WHERE attempt_id = ? AND question_id = ?
            """, (attempt_id, question_id))
            existing = cur.fetchone()
            
            if existing:
                cur.execute("""
                    UPDATE exam_attempt_answers 
                    SET selected_answer_id = coalesce(?, selected_answer_id),
                        answer_text = coalesce(?, answer_text)
                    WHERE attempt_id = ? AND question_id = ?
                """, (selected_answer_id, answer_text, attempt_id, question_id))
            else:
                cur.execute("""
                    INSERT INTO exam_attempt_answers (attempt_id, question_id, selected_answer_id, answer_text)
                    VALUES (?, ?, ?, ?)
                """, (attempt_id, question_id, selected_answer_id, answer_text))
            conn.commit()

    def get_saved_answers_for_attempt(self, attempt_id: int) -> dict:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT question_id, selected_answer_id FROM exam_attempt_answers WHERE attempt_id = ?", (attempt_id,))
            return {r['question_id']: r['selected_answer_id'] for r in cur.fetchall()}

    def get_saved_text_answers_for_attempt(self, attempt_id: int) -> dict:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(exam_attempt_answers)")
            cols = {r['name'] for r in cur.fetchall()}
            if 'answer_text' not in cols:
                return {}
            cur.execute("SELECT question_id, answer_text FROM exam_attempt_answers WHERE attempt_id = ? AND answer_text IS NOT NULL", (attempt_id,))
            return {r['question_id']: r['answer_text'] for r in cur.fetchall()}

    # Sequential level order — used for progression enforcement
    LEVEL_ORDER = ['Level 0', 'Level 1', 'Level 2', 'Level 3', 'Level 4']

    @classmethod
    def normalize_level(cls, lvl: str) -> str:
        if not lvl:
            return 'Level 0'
        s = str(lvl).strip()
        su = s.upper()
        if '100' in su or su == 'LEVEL 4' or su == 'LEVEL4' or 'EXPERT' in su or su == '4':
            return 'Level 4'
        if '75' in su or su == 'LEVEL 3' or su == 'LEVEL3' or su == '3':
            return 'Level 3'
        if '50' in su or su == 'LEVEL 2' or su == 'LEVEL2' or su == '2':
            return 'Level 2'
        if '25' in su or su == 'LEVEL 1' or su == 'LEVEL1' or su == '1':
            return 'Level 1'
        if '0' in su or 'BASE' in su or 'DÉBUTANT' in su or 'DEBUTANT' in su:
            return 'Level 0'
        # If already Title-cased "Level X"
        for l in cls.LEVEL_ORDER:
            if l.lower() == s.lower():
                return l
        return s

    def submit_exam_attempt(self, attempt_id: int):
        """
        Scores the attempt on the server, calculates pass/fail, saves metrics, and closes it.
        If passed, attempts to advance the technician level sequentially.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM exam_attempts WHERE id = ?", (attempt_id,))
            attempt = cur.fetchone()
            if not attempt or attempt['status'] != 'running':
                return False

            exam_id = attempt['exam_id']
            user_id = attempt['user_id']

            # Get exam specifications
            cur.execute("SELECT passing_score, number_of_questions, technician_level FROM exams WHERE id = ?", (exam_id,))
            exam_meta = cur.fetchone()
            passing_score_pct = exam_meta['passing_score']
            total_questions = exam_meta['number_of_questions']
            exam_target_level = exam_meta['technician_level']  # e.g. 'LEVEL 50%'

            # Get all questions
            cur.execute("SELECT id FROM exam_questions WHERE exam_id = ?", (exam_id,))
            q_ids = [r['id'] for r in cur.fetchall()]

            # Fetch selected answers & text answers
            cur.execute("PRAGMA table_info(exam_attempt_answers)")
            cols_ans = {r['name'] for r in cur.fetchall()}
            if 'answer_text' in cols_ans:
                cur.execute("SELECT question_id, selected_answer_id, answer_text FROM exam_attempt_answers WHERE attempt_id = ?", (attempt_id,))
                ans_rows = cur.fetchall()
                selected_map = {r['question_id']: r['selected_answer_id'] for r in ans_rows}
                text_map = {r['question_id']: r['answer_text'] for r in ans_rows}
            else:
                cur.execute("SELECT question_id, selected_answer_id FROM exam_attempt_answers WHERE attempt_id = ?", (attempt_id,))
                selected_map = {r['question_id']: r['selected_answer_id'] for r in cur.fetchall()}
                text_map = {}

            correct_count = 0
            incorrect_count = 0
            unanswered_count = 0

            # Only score auto-gradable question types
            _GRADABLE = ('multiple_choice', 'true_false', 'qcm', 'image_labeling', None)
            cur.execute(
                "SELECT id, question_type, boxes_json FROM exam_questions WHERE exam_id = ?", (exam_id,)
            )
            q_info_rows = cur.fetchall()
            q_type_map = {r['id']: r['question_type'] for r in q_info_rows}
            q_boxes_map = {r['id']: r['boxes_json'] for r in q_info_rows}

            gradable_q_ids = [
                qid for qid in q_ids
                if q_type_map.get(qid) in _GRADABLE
            ]

            import json as _json
            for q_id in gradable_q_ids:
                q_type = q_type_map.get(q_id)
                if q_type == 'image_labeling':
                    raw_text = text_map.get(q_id)
                    if not raw_text:
                        unanswered_count += 1
                        continue
                    try:
                        user_boxes = _json.loads(raw_text) if isinstance(raw_text, str) else raw_text
                    except Exception:
                        user_boxes = {}
                    
                    b_raw = q_boxes_map.get(q_id)
                    boxes_def = []
                    if b_raw:
                        try:
                            boxes_def = _json.loads(b_raw)
                        except Exception:
                            boxes_def = []
                    if not boxes_def:
                        boxes_def = DEFAULT_CSWIN_Q6_BOXES
                    
                    if not user_boxes:
                        unanswered_count += 1
                        continue
                    
                    # Grade box answers
                    box_correct = 0
                    for b in boxes_def:
                        b_id_str = str(b['id'])
                        target = (b.get('target') or '').strip().lower()
                        user_val = str(user_boxes.get(b_id_str) or '').strip().lower()
                        if target and user_val and (target == user_val or target in user_val or user_val in target):
                            box_correct += 1
                    
                    pass_threshold = max(1, int(len(boxes_def) * 0.7))
                    if box_correct >= pass_threshold:
                        correct_count += 1
                    else:
                        incorrect_count += 1
                else:
                    selected_ans_id = selected_map.get(q_id)
                    if not selected_ans_id:
                        unanswered_count += 1
                        continue

                    # Verify if this answer is correct in database
                    cur.execute("SELECT is_correct FROM exam_answers WHERE id = ? AND question_id = ?", (selected_ans_id, q_id))
                    ans_row = cur.fetchone()
                    if ans_row and ans_row['is_correct'] == 1:
                        correct_count += 1
                    else:
                        incorrect_count += 1

            pct = (correct_count / len(gradable_q_ids)) * 100 if gradable_q_ids else 0.0
            status = 'passed' if pct >= passing_score_pct else 'failed'

            cur.execute("""
                UPDATE exam_attempts 
                SET submitted_at = datetime('now', 'localtime'), 
                    score = ?, 
                    percentage = ?, 
                    correct_answers = ?, 
                    incorrect_answers = ?, 
                    unanswered = ?, 
                    status = ?
                WHERE id = ?
            """, (correct_count, round(pct, 1), correct_count, incorrect_count, unanswered_count, status, attempt_id))
            
            # ── SERVER-SIDE LEVEL ADVANCEMENT ────────────────────────────────
            # Get technician's current level
            cur.execute("SELECT name, matricule, role, shift, technician_level FROM users WHERE id = ?", (user_id,))
            user_row = cur.fetchone()
            level_advanced = False

            if status == 'passed' and user_row and user_row['role'] == 'TECHNICIAN':
                current_level = self.normalize_level(user_row['technician_level'])
                level_advanced = self._try_advance_level(
                    conn=conn,
                    cur=cur,
                    user_id=user_id,
                    current_level=current_level,
                    exam_target_level=self.normalize_level(exam_target_level),
                    attempt_id=attempt_id
                )

            # Sync to JSON history file if this user is a technician
            if user_row and user_row['role'] == 'TECHNICIAN' and user_row['matricule']:
                self._sync_attempt_to_json_profile(user_row['matricule'], exam_id, correct_count, total_questions, pct, status)
                
            conn.commit()
            return status

    def _try_advance_level(self, conn, cur, user_id: int, current_level: str, exam_target_level: str, attempt_id: int) -> bool:
        """
        Advances the technician level if:
        - They passed the exam
        - The exam's target level is EXACTLY one step above their current level
        Returns True if level was advanced, False otherwise.
        """
        lo = self.LEVEL_ORDER
        current_level = self.normalize_level(current_level)
        exam_target_level = self.normalize_level(exam_target_level)
        if current_level not in lo or exam_target_level not in lo:
            logger.warning(f"Level not in LEVEL_ORDER: current={current_level}, target={exam_target_level}")
            return False

        current_idx = lo.index(current_level)
        target_idx = lo.index(exam_target_level)

        # Only advance if target is EXACTLY one step ahead
        if target_idx != current_idx + 1:
            logger.info(
                f"Level NOT advanced: current={current_level}({current_idx}), "
                f"target={exam_target_level}({target_idx}) — not sequential next step."
            )
            return False

        # Advance the level
        new_level = exam_target_level
        cur.execute(
            "UPDATE users SET technician_level = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (new_level, user_id)
        )

        # Record in audit log
        cur.execute("""
            INSERT INTO level_change_log 
                (user_id, previous_level, new_level, changed_by, change_type, reason, exam_attempt_id)
            VALUES (?, ?, ?, 'SYSTEM', 'EXAM', ?, ?)
        """, (user_id, current_level, new_level, f'Examen passé avec succès — niveau {new_level} débloqué.', attempt_id))

        logger.info(f"Technician user_id={user_id} advanced: {current_level} → {new_level}")
        return True

    def _sync_attempt_to_json_profile(self, matricule, exam_id, score, total_qs, pct, status):
        safe_mat = str(matricule).replace("/", "_").replace("\\", "_").strip()
        p = os.path.join(self.data_dir, "Technicians", f"{safe_mat}.json")
        if os.path.exists(p):
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    profile = json.load(f)
                
                # Fetch exam title
                with self._get_conn() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT title, technician_level FROM exams WHERE id = ?", (exam_id,))
                    ex = cur.fetchone()
                    exam_title = ex['title'] if ex else "Examen"
                    exam_level = ex['technician_level'] if ex else ""

                if 'exams' not in profile:
                    profile['exams'] = []

                # Remove existing run of same exam to show latest
                profile['exams'] = [e for e in profile['exams'] if e.get('exam_id') != exam_id]

                profile['exams'].append({
                    "exam_id": exam_id,
                    "name": exam_title,
                    "type": "Exam",
                    "exam_level": exam_level,
                    "date": datetime.date.today().isoformat(),
                    "score": f"{score}/{total_qs}",
                    "percentage": pct,
                    "status": "RÉUSSI" if status == 'passed' else "ÉCHOUÉ"
                })
                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(profile, f, indent=2, ensure_ascii=False)
            except Exception as e:
                logger.error(f"Failed to sync attempt to JSON profile {matricule}: {e}")

    # ── Level Progression Helpers ─────────────────────────────────────────

    def get_technician_progression(self, user_id: int) -> dict:
        """
        Returns a structured progression object for the 4 levels:
        - For each level: status ('current', 'unlocked', 'locked'), exam attempt if any.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name, matricule, technician_level FROM users WHERE id = ?", (user_id,))
            user_row = cur.fetchone()
            if not user_row:
                return {}

            current_level = self.normalize_level(user_row['technician_level'])
            lo = self.LEVEL_ORDER

            if current_level not in lo:
                current_idx = 0
            else:
                current_idx = lo.index(current_level)

            level_descriptions = {
                'Level 0': 'Niveau Initial (Débutant)',
                'Level 1': 'CSwin Basic Knowledge',
                'Level 2': 'CSwin Creation Hardware',
                'Level 3': 'CSwin & Brainware & Vacuum',
                'Level 4': 'CSwin & Brainware & Vacuum',
                'LEVEL 0%':   'Niveau Initial (Débutant)',
                'LEVEL 25%':  'CSwin Basic Knowledge',
                'LEVEL 50%':  'CSwin Creation Hardware',
                'LEVEL 75%':  'CSwin & Brainware & Vacuum',
                'LEVEL 100%': 'CSwin & Brainware & Vacuum',
            }

            levels = []
            for idx, lvl in enumerate(lo):
                if idx < current_idx:
                    status = 'unlocked'
                elif idx == current_idx:
                    status = 'current'
                else:
                    status = 'locked'

                # Get best passing attempt for this level's exam
                cur.execute("""
                    SELECT ea.id, ea.percentage, ea.submitted_at, ea.status as attempt_status, e.title as exam_title
                    FROM exam_attempts ea
                    JOIN exams e ON ea.exam_id = e.id
                    WHERE ea.user_id = ? AND e.technician_level = ? AND ea.status = 'passed'
                    ORDER BY ea.submitted_at DESC LIMIT 1
                """, (user_id, lvl))
                best_attempt = cur.fetchone()

                # Also get latest attempt regardless of pass/fail
                cur.execute("""
                    SELECT ea.id, ea.percentage, ea.submitted_at, ea.status as attempt_status, e.title as exam_title
                    FROM exam_attempts ea
                    JOIN exams e ON ea.exam_id = e.id
                    WHERE ea.user_id = ? AND e.technician_level = ? AND ea.status != 'running'
                    ORDER BY ea.submitted_at DESC LIMIT 1
                """, (user_id, lvl))
                latest_attempt = cur.fetchone()

                # Fetch published exam for this level
                cur.execute("""
                    SELECT id, title, duration, number_of_questions, passing_score
                    FROM exams WHERE technician_level = ? AND status = 'published'
                    LIMIT 1
                """, (lvl,))
                ex_row = cur.fetchone()

                levels.append({
                    'level': lvl,
                    'description': level_descriptions.get(lvl, ''),
                    'status': status,
                    'exam': dict(ex_row) if ex_row else None,
                    'best_attempt': dict(best_attempt) if best_attempt else None,
                    'latest_attempt': dict(latest_attempt) if latest_attempt else None,
                })

            next_lvl = lo[current_idx + 1] if current_idx + 1 < len(lo) else None
            next_exam = None
            if next_lvl:
                cur.execute("""
                    SELECT id, title, duration, number_of_questions, passing_score
                    FROM exams WHERE technician_level = ? AND status = 'published'
                    LIMIT 1
                """, (next_lvl,))
                ne_row = cur.fetchone()
                if ne_row:
                    next_exam = dict(ne_row)

            return {
                'user': dict(user_row),
                'current_level': current_level,
                'current_description': level_descriptions.get(current_level, ''),
                'next_level': next_lvl,
                'next_exam': next_exam,
                'levels': levels,
            }

    def get_technician_level_history(self, user_id: int) -> list:
        """
        Returns full audit history of level changes for a technician.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT lcl.*, ea.percentage as exam_percentage
                FROM level_change_log lcl
                LEFT JOIN exam_attempts ea ON lcl.exam_attempt_id = ea.id
                WHERE lcl.user_id = ?
                ORDER BY lcl.changed_at DESC
            """, (user_id,))
            return [dict(r) for r in cur.fetchall()]

    def owner_override_level(
        self,
        user_id: int,
        new_level: str,
        reason: str,
        owner_username: str,
        create_cert: bool = False
    ) -> tuple:
        """
        Allows OWNER to manually override a technician's level.
        Records the change in level_change_log.
        If create_cert is True, also creates a passed exam attempt record so the technician's card is certified.
        Returns (success: bool, message: str).
        """
        target_lvl = self.normalize_level(new_level)
        if target_lvl not in self.LEVEL_ORDER:
            return False, f"Niveau invalide : {new_level}"
        new_level = target_lvl

        if not reason or len(reason.strip()) < 5:
            return False, "Une raison d'au moins 5 caractères est obligatoire pour une modification manuelle."

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name, technician_level, role FROM users WHERE id = ?", (user_id,))
            user_row = cur.fetchone()
            if not user_row:
                return False, "Technicien introuvable."

            if user_row['role'] != 'TECHNICIAN':
                return False, "Cette action ne s'applique qu'aux techniciens."

            previous_level = self.normalize_level(user_row['technician_level'])

            if previous_level == new_level:
                return False, f"Le technicien est déjà au niveau {new_level}."

            cur.execute(
                "UPDATE users SET technician_level = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                (new_level, user_id)
            )

            exam_attempt_id = None
            if create_cert:
                # Find the exam for this level
                cur.execute("SELECT id, number_of_questions FROM exams WHERE technician_level = ? ORDER BY id DESC LIMIT 1", (new_level,))
                ex_row = cur.fetchone()
                if ex_row:
                    ex_id = ex_row['id']
                    num_q = ex_row['number_of_questions'] or 10
                    cur.execute("""
                        INSERT INTO exam_attempts (exam_id, user_id, started_at, submitted_at, score, percentage, correct_answers, incorrect_answers, unanswered, status)
                        VALUES (?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, ?, 100.0, ?, 0, 0, 'passed')
                    """, (ex_id, user_id, num_q, num_q))
                    exam_attempt_id = cur.lastrowid

            cur.execute("""
                INSERT INTO level_change_log 
                    (user_id, previous_level, new_level, changed_by, change_type, reason, exam_attempt_id)
                VALUES (?, ?, ?, ?, 'MANUAL_OVERRIDE', ?, ?)
            """, (user_id, previous_level, new_level, owner_username, reason.strip(), exam_attempt_id))
            conn.commit()

            logger.info(
                f"OWNER OVERRIDE: user_id={user_id} ({user_row['name']}) "
                f"{previous_level} → {new_level} by {owner_username}. Reason: {reason} (Cert: {create_cert})"
            )
            cert_msg = " avec attestation d'examen certifiée" if create_cert else ""
            return True, f"Niveau de {user_row['name']} modifié manuellement : {previous_level} → {new_level}{cert_msg}."

    def get_all_technicians_with_progression(self) -> list:
        """Returns all technicians with their current progression state."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, name, matricule, shift, technician_level, photo
                FROM users WHERE role = 'TECHNICIAN'
                ORDER BY name ASC
            """)
            techs = [dict(r) for r in cur.fetchall()]

        level_descriptions = {
            'Level 0': 'Niveau Initial (Débutant)',
            'Level 1': 'CSwin Basic Knowledge',
            'Level 2': 'CSwin Creation Hardware',
            'Level 3': 'CSwin & Brainware & Vacuum',
            'Level 4': 'CSwin & Brainware & Vacuum',
            'LEVEL 0%':   'Niveau Initial (Débutant)',
            'LEVEL 25%':  'CSwin Basic Knowledge',
            'LEVEL 50%':  'CSwin Creation Hardware',
            'LEVEL 75%':  'CSwin & Brainware & Vacuum',
            'LEVEL 100%': 'CSwin & Brainware & Vacuum',
        }
        for t in techs:
            lvl = self.normalize_level(t.get('technician_level'))
            t['technician_level'] = lvl
            t['level_description'] = level_descriptions.get(lvl, '')
            lo = self.LEVEL_ORDER
            t['level_idx'] = lo.index(lvl) if lvl in lo else 0
            t['level_pct'] = t['level_idx'] * 25
        return techs

    # ── Multi-Image Helpers ────────────────────────────────────────────────────

    def _get_question_extra_images_cur(self, cur, question_id: int) -> list:
        """Internal: fetch extra images using an existing cursor."""
        cur.execute(
            "SELECT * FROM exam_question_images WHERE question_id = ? ORDER BY sort_order ASC, id ASC",
            (question_id,)
        )
        return [dict(r) for r in cur.fetchall()]

    def get_question_extra_images(self, question_id: int) -> list:
        """Returns all rows from exam_question_images for a question."""
        with self._get_conn() as conn:
            return self._get_question_extra_images_cur(conn.cursor(), question_id)

    def add_question_extra_image(self, question_id: int, filename: str, sort_order: int = 0) -> int:
        """Adds a filename to exam_question_images for a question. Returns new row id."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO exam_question_images (question_id, filename, sort_order) VALUES (?, ?, ?)",
                (question_id, filename, sort_order)
            )
            conn.commit()
            return cur.lastrowid

    def remove_extra_image(self, image_row_id: int) -> bool:
        """Removes a specific extra image row."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM exam_question_images WHERE id = ?", (image_row_id,))
            conn.commit()
            return True

    def move_question_image(self, source_question_id: int, target_question_id: int) -> tuple:
        """
        Moves the primary image of source_question_id to target_question_id.
        - Sets source question's image = NULL
        - Sets target question's image = the filename (overwriting any existing)
        - Also migrates the filename in exam_question_images if present.
        Returns (success: bool, message: str).
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, exam_id, image FROM exam_questions WHERE id = ?", (source_question_id,))
            src = cur.fetchone()
            if not src:
                return False, "Question source introuvable."
            if not src['image']:
                return False, "La question source n'a pas d'image à déplacer."

            cur.execute("SELECT id, exam_id FROM exam_questions WHERE id = ?", (target_question_id,))
            tgt = cur.fetchone()
            if not tgt:
                return False, "Question cible introuvable."

            if src['exam_id'] != tgt['exam_id']:
                return False, "Les deux questions doivent appartenir au même examen."

            filename = src['image']

            # Remove image from source
            cur.execute("UPDATE exam_questions SET image = NULL WHERE id = ?", (source_question_id,))
            # Assign to target
            cur.execute("UPDATE exam_questions SET image = ? WHERE id = ?", (filename, target_question_id))
            # Migrate exam_question_images rows as well
            cur.execute(
                "UPDATE exam_question_images SET question_id = ? WHERE question_id = ? AND filename = ?",
                (target_question_id, source_question_id, filename)
            )
            conn.commit()
            return True, f"Image '{filename}' déplacée vers la question {target_question_id}."

    def remove_question_image(self, question_id: int) -> bool:
        """Removes the primary image and all extra images of a question."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE exam_questions SET image = NULL WHERE id = ?", (question_id,))
            cur.execute("DELETE FROM exam_question_images WHERE question_id = ?", (question_id,))
            conn.commit()
            return True

    # ── Admin Exam Editing Helpers ────────────────────────────────────────────

    def update_question_image(self, question_id: int, filename: str):
        """Sets the image filename for a question."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE exam_questions SET image = ? WHERE id = ?", (filename, question_id))
            conn.commit()
            return True

    def update_question_text(self, question_id: int, question_text: str):
        """Updates the text of a question."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE exam_questions SET question_text = ? WHERE id = ?", (question_text.strip(), question_id))
            conn.commit()
            return True

    def add_question_with_choices(self, exam_id: int, question_text: str, choices: list) -> int:
        """Adds a new question with choices to an exam. choices = [{'answer_text': ..., 'is_correct': 0/1}]."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT MAX(question_number) as mx FROM exam_questions WHERE exam_id = ?", (exam_id,))
            row = cur.fetchone()
            next_num = (row['mx'] or 0) + 1
            cur.execute(
                "INSERT INTO exam_questions (exam_id, question_number, question_text, question_type) VALUES (?, ?, ?, 'qcm')",
                (exam_id, next_num, question_text.strip())
            )
            q_id = cur.lastrowid
            for c in choices:
                cur.execute(
                    "INSERT INTO exam_answers (question_id, answer_text, is_correct) VALUES (?, ?, ?)",
                    (q_id, c['answer_text'], c.get('is_correct', 0))
                )
            cur.execute(
                "UPDATE exams SET number_of_questions = (SELECT COUNT(*) FROM exam_questions WHERE exam_id = ?), updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                (exam_id, exam_id)
            )
            conn.commit()
            return q_id

    def delete_question(self, question_id: int):
        """Deletes a question and all its answers (cascade)."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT exam_id FROM exam_questions WHERE id = ?", (question_id,))
            row = cur.fetchone()
            if not row:
                return False
            exam_id = row['exam_id']
            cur.execute("DELETE FROM exam_questions WHERE id = ?", (question_id,))
            # Renumber remaining questions
            cur.execute("SELECT id FROM exam_questions WHERE exam_id = ? ORDER BY question_number ASC", (exam_id,))
            remaining = cur.fetchall()
            for idx, r in enumerate(remaining):
                cur.execute("UPDATE exam_questions SET question_number = ? WHERE id = ?", (idx + 1, r['id']))
            cur.execute(
                "UPDATE exams SET number_of_questions = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                (len(remaining), exam_id)
            )
            conn.commit()
            return True

    def toggle_publish_exam(self, exam_id: int) -> str:
        """Toggles exam status between 'published' and 'draft'."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT status FROM exams WHERE id = ?", (exam_id,))
            row = cur.fetchone()
            if not row:
                return 'not_found'
            new_status = 'draft' if row['status'] == 'published' else 'published'
            cur.execute("UPDATE exams SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_status, exam_id))
            conn.commit()
            return new_status

    def create_blank_exam(self, title: str, technician_level: str, level_percentage: int) -> int:
        """Creates a blank draft exam for the given level."""
        technician_level = self.normalize_level(technician_level)
        level_descriptions = {
            'Level 0': 'Niveau Initial (Débutant)',
            'Level 1': 'CSwin Basic Knowledge',
            'Level 2': 'CSwin Creation Hardware',
            'Level 3': 'CSwin & Brainware & Vacuum',
            'Level 4': 'CSwin & Brainware & Vacuum',
            'LEVEL 0%':   'Niveau Initial (Débutant)',
            'LEVEL 25%':  'CSwin Basic Knowledge',
            'LEVEL 50%':  'CSwin Creation Hardware',
            'LEVEL 75%':  'CSwin & Brainware & Vacuum',
            'LEVEL 100%': 'CSwin & Brainware & Vacuum',
        }
        desc = level_descriptions.get(technician_level, '')
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO exams (title, description, technician_level, level_percentage, duration, number_of_questions, passing_score, status)
                VALUES (?, ?, ?, ?, 30, 0, 75, 'draft')
            """, (title, desc, technician_level, level_percentage))
            conn.commit()
            return cur.lastrowid

    def get_all_technicians_for_override(self) -> list:
        """Returns all technicians for the manual level override dropdown."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name, matricule, shift, technician_level FROM users WHERE role = 'TECHNICIAN' ORDER BY name ASC")
            return [dict(r) for r in cur.fetchall()]

    # ── History & Technician Card Data Retrieval ────────────────────────────

    def get_all_results(self, filters=None):
        query = """
            SELECT ea.id as attempt_id, ea.started_at, ea.submitted_at, ea.score, ea.percentage, ea.status,
                   u.name as tech_name, u.matricule as tech_mat, u.technician_level,
                   e.title as exam_title, e.number_of_questions
            FROM exam_attempts ea
            JOIN users u ON ea.user_id = u.id
            JOIN exams e ON ea.exam_id = e.id
            WHERE ea.status != 'running'
        """
        params = []
        if filters:
            if filters.get('technician'):
                query += " AND u.name LIKE ?"
                params.append(f"%{filters['technician']}%")
            if filters.get('matricule'):
                query += " AND u.matricule LIKE ?"
                params.append(f"%{filters['matricule']}%")
            if filters.get('level'):
                query += " AND u.technician_level = ?"
                params.append(filters['level'])
            if filters.get('exam_id'):
                query += " AND ea.exam_id = ?"
                params.append(filters['exam_id'])
            if filters.get('status'):
                query += " AND ea.status = ?"
                params.append(filters['status'])
            if filters.get('date'):
                query += " AND ea.submitted_at LIKE ?"
                params.append(f"{filters['date']}%")

        query += " ORDER BY ea.submitted_at DESC"
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(r) for r in cur.fetchall()]

    def get_technician_card_details(self, user_id: int):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name, matricule, shift, technician_level, photo FROM users WHERE id = ?", (user_id,))
            user_row = cur.fetchone()
            if not user_row:
                return None
                
            tech = dict(user_row)
            tech['technician_level'] = self.normalize_level(tech.get('technician_level'))
            tech['shift'] = tech.get('shift') or 'A'
            
            cur.execute("""
                SELECT ea.*, e.title as exam_title, e.number_of_questions
                FROM exam_attempts ea
                JOIN exams e ON ea.exam_id = e.id
                WHERE ea.user_id = ? AND ea.status != 'running'
                ORDER BY ea.submitted_at DESC LIMIT 1
            """, (user_id,))
            attempt_row = cur.fetchone()
            if attempt_row:
                tech['has_exam'] = True
                tech['exam_title'] = attempt_row['exam_title']
                tech['score'] = f"{attempt_row['score']}/{attempt_row['number_of_questions']}"
                tech['percentage'] = attempt_row['percentage']
                tech['status'] = "RÉUSSI" if attempt_row['status'] == 'passed' else "ÉCHOUÉ"
                tech['date'] = attempt_row['submitted_at'][:10] if attempt_row['submitted_at'] else ""
                tech['card_id'] = f"CERT-{attempt_row['id']:05d}"
            else:
                tech['has_exam'] = False
                tech['exam_title'] = "Aucun examen passé"
                tech['score'] = "N/A"
                tech['percentage'] = 0.0
                tech['status'] = "Nouveau"
                tech['date'] = "-"
                tech['card_id'] = f"TECH-{tech['id']:05d}"
                
            return tech

    # ── Question Type & Available Labels Helpers ─────────────────────────────

    def set_question_type(self, question_id: int, new_type: str) -> bool:
        """Sets the question_type for a question ('multiple_choice', 'image_labeling', 'true_false', 'text_input')."""
        valid_types = ('multiple_choice', 'image_labeling', 'true_false', 'text_input')
        if new_type not in valid_types:
            return False
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE exam_questions SET question_type = ? WHERE id = ?", (new_type, question_id))
            conn.commit()
            return True

    def set_available_labels(self, question_id: int, labels: list) -> bool:
        """Sets the available_labels JSON array for an image_labeling question."""
        labels_json = json.dumps(labels, ensure_ascii=False) if labels else None
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE exam_questions SET available_labels = ? WHERE id = ?", (labels_json, question_id))
            conn.commit()
            return True

    def fix_question_types_for_exam(self, exam_id: int) -> dict:
        """
        Retroactively analyzes and updates question types and available labels for an exam.
        Detects image_labeling, true_false, multiple_choice, and text_input.
        """
        updated_count = 0
        details = []

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, question_number, question_text, question_type, image, available_labels FROM exam_questions WHERE exam_id = ? ORDER BY question_number", (exam_id,))
            questions = [dict(r) for r in cur.fetchall()]

            for q in questions:
                cur.execute("SELECT id, answer_text, is_correct FROM exam_answers WHERE question_id = ?", (q['id'],))
                answers = [dict(a) for a in cur.fetchall()]

                text_low = q['question_text'].lower()
                has_labeling_kw = any(kw in text_low for kw in self._LABELING_KEYWORDS)
                has_image = bool(q['image'])

                new_type = q.get('question_type') or 'multiple_choice'
                new_labels = None

                if has_labeling_kw and (has_image or q.get('available_labels') or answers):
                    new_type = 'image_labeling'
                    # If answers exist in exam_answers, move them to available_labels
                    if answers and not q.get('available_labels'):
                        label_list = [a['answer_text'] for a in answers if a['answer_text'].strip()]
                        new_labels = json.dumps(label_list, ensure_ascii=False)
                        # Remove answers from exam_answers table since it's not a QCM
                        cur.execute("DELETE FROM exam_answers WHERE question_id = ?", (q['id'],))
                elif len(answers) == 2:
                    texts = {a['answer_text'].strip().lower() for a in answers}
                    tf_pairs = [
                        {'vrai', 'faux'}, {'true', 'false'},
                        {'oui', 'non'}, {'yes', 'no'},
                        {'correct', 'incorrect'}, {'juste', 'faux'},
                    ]
                    if any(texts == pair for pair in tf_pairs):
                        new_type = 'true_false'
                    else:
                        new_type = 'multiple_choice'
                elif answers:
                    new_type = 'multiple_choice'
                else:
                    new_type = 'text_input'

                if new_type != q.get('question_type') or new_labels is not None:
                    if new_labels is not None:
                        cur.execute("UPDATE exam_questions SET question_type = ?, available_labels = ? WHERE id = ?", (new_type, new_labels, q['id']))
                    else:
                        cur.execute("UPDATE exam_questions SET question_type = ? WHERE id = ?", (new_type, q['id']))
                    updated_count += 1
                    details.append({
                        'question_id': q['id'],
                        'question_number': q['question_number'],
                        'old_type': q.get('question_type'),
                        'new_type': new_type
                    })

            conn.commit()

        return {'exam_id': exam_id, 'updated': updated_count, 'details': details}

    def fix_all_question_types(self) -> dict:
        """Runs retroactive question type detection across all exams."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id FROM exams")
            exam_ids = [r['id'] for r in cur.fetchall()]

        total_updated = 0
        all_results = []
        for eid in exam_ids:
            res = self.fix_question_types_for_exam(eid)
            total_updated += res['updated']
            all_results.append(res)

        return {'total_updated': total_updated, 'exams': all_results}

    def reimport_official_exams(self) -> dict:
        """
        Re-parses the official test docx files in downloads and updates/replaces
        existing official exams with full question type detection and image extraction.
        """
        download_dir = r"C:\Users\hama0\Downloads"
        if not os.path.isdir(download_dir):
            return {'success': False, 'message': 'Dossier de téléchargement introuvable'}

        levels_map = {
            "25%": {"title": "CSwin Basic Knowledge", "pct": 25, "level": "Level 1"},
            "50%": {"title": "CSwin Creation Hardware", "pct": 50, "level": "Level 2"},
            "75%": {"title": "CSwin & Brainware & Vacuum", "pct": 75, "level": "Level 3"},
            "100%": {"title": "CSwin & Brainware & Vacuum", "pct": 100, "level": "Level 4"}
        }

        img_out = os.path.join(self.data_dir, "exam_images")
        alt_img = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "web_portal", "static", "exam_images"))
        if os.path.exists(alt_img):
            img_out = alt_img
        os.makedirs(img_out, exist_ok=True)

        try:
            files = os.listdir(download_dir)
        except Exception as e:
            return {'success': False, 'error': str(e)}

        reimported = []

        for f in files:
            if "Test Level" in f and f.endswith(".docx"):
                lvl_key = None
                for key in levels_map:
                    if key in f:
                        lvl_key = key
                        break

                if lvl_key:
                    cfg = levels_map[lvl_key]
                    full_path = os.path.join(download_dir, f)
                    try:
                        existing_id = None
                        with self._get_conn() as conn:
                            cur = conn.cursor()
                            cur.execute("SELECT id, status FROM exams WHERE source_file = ?", (f,))
                            row = cur.fetchone()
                            if row:
                                existing_id = row['id']

                        questions = self.parse_docx_exam(full_path, output_images_dir=img_out)

                        if existing_id:
                            with self._get_conn() as conn:
                                cur = conn.cursor()
                                cur.execute("DELETE FROM exam_questions WHERE exam_id = ?", (existing_id,))
                                cur.execute("UPDATE exams SET number_of_questions = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (len(questions), existing_id))
                                for q in questions:
                                    q_type = q.get('question_type', 'multiple_choice')
                                    labels_json = json.dumps(q.get('available_labels', []), ensure_ascii=False) if q.get('available_labels') else None
                                    cur.execute("""
                                        INSERT INTO exam_questions (exam_id, question_number, question_text, question_type, image, available_labels)
                                        VALUES (?, ?, ?, ?, ?, ?)
                                    """, (existing_id, q['question_number'], q['question_text'], q_type, q.get('image'), labels_json))
                                    q_id = cur.lastrowid
                                    for sort_i, ef in enumerate(q.get('extra_images', [])):
                                        cur.execute("INSERT INTO exam_question_images (question_id, filename, sort_order) VALUES (?, ?, ?)", (q_id, ef, sort_i + 1))
                                    if q_type in ('multiple_choice', 'true_false', 'qcm'):
                                        for choice in q.get('choices', []):
                                            cur.execute("INSERT INTO exam_answers (question_id, answer_text, is_correct) VALUES (?, ?, ?)", (q_id, choice['answer_text'], choice.get('is_correct', 0)))
                                conn.commit()
                            reimported.append({'file': f, 'level': cfg['level'], 'id': existing_id, 'questions': len(questions)})
                        else:
                            new_id = self.create_exam_with_questions(
                                title=cfg["title"],
                                description=f"Examen officiel pour le niveau {cfg['level']}",
                                technician_level=cfg["level"],
                                level_percentage=cfg["pct"],
                                source_file=f,
                                duration=30,
                                passing_score=75,
                                status='published',
                                parsed_questions=questions
                            )
                            reimported.append({'file': f, 'level': cfg['level'], 'id': new_id, 'questions': len(questions)})
                    except Exception as e:
                        logger.error(f"Error reimporting {f}: {e}")

        return {'success': True, 'reimported': reimported}


