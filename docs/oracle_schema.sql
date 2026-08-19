-- =============================================================================
-- טיוטת סכימה יחסית ל-Oracle — מערכת מבחני תאוריה
-- =============================================================================
-- מקור: תרגום 14 לשוניות ה-Google Sheets (ראה DATA_MODEL.md) למודל יחסי,
-- כולל תיקון הפערים המבניים של הגיליון (ראה הערות "שיפור" בהמשך).
-- זו טיוטה לדיון עם צוות הפיתוח — שמות/טיפוסים ניתנים להתאמה למוסכמות שלכם.
-- תחביר: Oracle 19c ומעלה (IDENTITY, VARCHAR2, CLOB).
--
-- שלושה עקרונות שאסור לאבד בתרגום (מוסבר ב-FUNCTIONAL_SPEC.md):
--   1. ת.ז. היא VARCHAR2(9) עם אפסים מובילים — לעולם לא NUMBER.
--   2. תוצאות לא נמחקות — רק מסומנות (status/flags).
--   3. "התשובה הנכונה" מוגדרת פר-שפה (סדר תשובות שונה בין תרגומים).
-- =============================================================================

-- ---------------------------------------------------------------------------
-- 1. משתמשים ואתרים
-- ---------------------------------------------------------------------------

CREATE TABLE sites (
  site_id        NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  site_name      VARCHAR2(120 CHAR) NOT NULL UNIQUE,
  manager_phone  VARCHAR2(20 CHAR),
  is_test_site   NUMBER(1) DEFAULT 0 NOT NULL,  -- אתרי בדיקה מוחרגים מסטטיסטיקות
  is_active      NUMBER(1) DEFAULT 1 NOT NULL
);
COMMENT ON TABLE sites IS 'אתרים/בסיסים. is_test_site מחליף את רשימת TEST_SITES הקשיחה';

CREATE TABLE examiners (
  examiner_id    NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  national_id    VARCHAR2(9 CHAR) NOT NULL UNIQUE,   -- ת.ז. מנורמלת, 9 ספרות עם אפסים מובילים
  full_name      VARCHAR2(80 CHAR) NOT NULL,
  password_hash  VARCHAR2(255 CHAR) NOT NULL,        -- שיפור: היום סיסמה גלויה בגיליון
  role_code      VARCHAR2(20 CHAR) DEFAULT 'EXAMINER' NOT NULL
                 CHECK (role_code IN ('EXAMINER','SENIOR_EXAMINER','COMMANDER',
                                      'LOCAL_COMMANDER','CENTER_COMMANDER','CHIEF_COMMANDER','ADMIN')),
  is_active      NUMBER(1) DEFAULT 1 NOT NULL,
  failed_logins  NUMBER(3) DEFAULT 0 NOT NULL,
  locked_until   TIMESTAMP
);

-- אתרים מנוהלים למפקד-מרכז (היום: מחרוזת מופרדת בפסיקים בגיליון)
CREATE TABLE examiner_managed_sites (
  examiner_id  NUMBER NOT NULL REFERENCES examiners(examiner_id),
  site_id      NUMBER NOT NULL REFERENCES sites(site_id),
  PRIMARY KEY (examiner_id, site_id)
);

CREATE TABLE teachers (
  teacher_id     NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  national_id    VARCHAR2(9 CHAR) NOT NULL UNIQUE,
  full_name      VARCHAR2(80 CHAR) NOT NULL,
  password_hash  VARCHAR2(255 CHAR) NOT NULL,
  role_code      VARCHAR2(20 CHAR) DEFAULT 'TEACHER' NOT NULL
                 CHECK (role_code IN ('TEACHER','TEACHER_COMMANDER')),
  site_id        NUMBER REFERENCES sites(site_id),
  is_active      NUMBER(1) DEFAULT 1 NOT NULL,
  failed_logins  NUMBER(3) DEFAULT 0 NOT NULL,
  locked_until   TIMESTAMP
);

-- ---------------------------------------------------------------------------
-- 2. בנק השאלות (7 שפות; הנכונות פר-שפה!)
-- ---------------------------------------------------------------------------

CREATE TABLE questions (
  question_id  NUMBER PRIMARY KEY,                    -- לשמר את המזהים המקוריים! (ייבוא)
  category     VARCHAR2(40 CHAR) NOT NULL,            -- בטיחות / הכרת הרכב / חוק / תמרורים / ספציפי-*
  image_url    VARCHAR2(500 CHAR),
  is_active    NUMBER(1) DEFAULT 1 NOT NULL           -- הסרת שאלה = כיבוי, לא מחיקה
);

CREATE TABLE question_licenses (
  question_id  NUMBER NOT NULL REFERENCES questions(question_id),
  license_code VARCHAR2(3 CHAR) NOT NULL CHECK (license_code IN ('B','1','C1','C','D')),
  PRIMARY KEY (question_id, license_code)
);
COMMENT ON TABLE question_licenses IS 'שאלה משרתת כמה דרגות (M:N)';

CREATE TABLE question_texts (
  question_id   NUMBER NOT NULL REFERENCES questions(question_id),
  lang_code     VARCHAR2(2 CHAR) NOT NULL CHECK (lang_code IN ('he','ru','ar','am','en','fr','es')),
  question_text CLOB NOT NULL,
  PRIMARY KEY (question_id, lang_code)
);

CREATE TABLE question_answers (
  question_id  NUMBER NOT NULL,
  lang_code    VARCHAR2(2 CHAR) NOT NULL,
  position     NUMBER(1) NOT NULL,                    -- 0-3, הסדר הקנוני של אותה שפה
  answer_text  VARCHAR2(1000 CHAR) NOT NULL,
  is_correct   NUMBER(1) DEFAULT 0 NOT NULL,
  PRIMARY KEY (question_id, lang_code, position),
  FOREIGN KEY (question_id, lang_code) REFERENCES question_texts(question_id, lang_code)
);
-- אכיפה: בדיוק תשובה נכונה אחת לכל (שאלה, שפה)
CREATE UNIQUE INDEX ux_qa_one_correct
  ON question_answers (question_id, lang_code, CASE WHEN is_correct = 1 THEN 1 ELSE NULL END, CASE WHEN is_correct = 1 THEN NULL ELSE ROWNUM END);
-- ^ אם המבנה הזה מסורבל אצלכם — טריגר או materialized check; העיקרון: אחת ויחידה.
COMMENT ON TABLE question_answers IS 'זהו מפתח התשובות. סדר התשובות שונה בין שפות — is_correct פר-שפה. סודי: אסור שיגיע ללקוח בזמן מבחן';

-- מבנה המבחן (blueprint) — כמה שאלות מכל נושא לכל דרגה
CREATE TABLE exam_blueprint (
  license_code VARCHAR2(3 CHAR) NOT NULL,
  category     VARCHAR2(40 CHAR) NOT NULL,
  q_count      NUMBER(2) NOT NULL,
  PRIMARY KEY (license_code, category)
);
COMMENT ON TABLE exam_blueprint IS 'B:7/7/7/9 · 1:5/5/6/6/8 · C1:5/5/5/5/10 · C:5/4/3/4/14 · D:4/2/5/4/15 — סה"כ 30. דרישה תוכנית, שינוי רק באישור גורם מקצועי';

-- ---------------------------------------------------------------------------
-- 3. סשנים ורישומי נבחנים
-- ---------------------------------------------------------------------------

CREATE TABLE exam_sessions (
  session_id       NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  session_code     VARCHAR2(8 CHAR) NOT NULL UNIQUE,  -- הקוד שהנבחן מקליד
  examiner_id      NUMBER NOT NULL REFERENCES examiners(examiner_id),  -- הבעלים; רק הוא מעדכן
  host_site_id     NUMBER NOT NULL REFERENCES sites(site_id),
  classroom        VARCHAR2(80 CHAR),
  license_code     VARCHAR2(3 CHAR) NOT NULL,
  default_lang     VARCHAR2(2 CHAR) DEFAULT 'he' NOT NULL,
  default_audio    NUMBER(1) DEFAULT 0 NOT NULL,      -- ברירת מחדל לסימון-מראש בלבד (שמע אמיתי = פר-רישום)
  default_population VARCHAR2(40 CHAR),
  responsible_examiner_id NUMBER REFERENCES examiners(examiner_id),   -- נדרש לדו"ח משותף
  created_at       TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
  valid_until      TIMESTAMP NOT NULL,                -- יצירה + 8 שעות
  is_active        NUMBER(1) DEFAULT 1 NOT NULL,
  report_shared_at TIMESTAMP                          -- סגירה חסומה עד שזה מלא (אם יש תוצאות)
);

-- אתרי אורח: סשן מארח נבחנים מכמה בסיסים
CREATE TABLE session_guest_sites (
  session_id  NUMBER NOT NULL REFERENCES exam_sessions(session_id),
  site_id     NUMBER NOT NULL REFERENCES sites(site_id),
  PRIMARY KEY (session_id, site_id)
);

CREATE TABLE exam_registrations (
  registration_id  NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  session_id       NUMBER NOT NULL REFERENCES exam_sessions(session_id),
  national_id      VARCHAR2(9 CHAR) NOT NULL,
  full_name        VARCHAR2(80 CHAR) NOT NULL,
  phone            VARCHAR2(20 CHAR),
  registered_at    TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
  status           VARCHAR2(15 CHAR) DEFAULT 'WAITING' NOT NULL
                   CHECK (status IN ('WAITING','APPROVED','IN_EXAM','COMPLETED',
                                     'DISQUALIFIED','DQ_CONFIRMED','REJECTED','CANCELLED')),
  lang_code        VARCHAR2(2 CHAR) NOT NULL,
  population       VARCHAR2(40 CHAR),
  license_code     VARCHAR2(3 CHAR) NOT NULL,          -- בחירת הנבחן; יכולה לחרוג מדרגת הסשן
  chosen_site_id   NUMBER REFERENCES sites(site_id),   -- מארח/אורח
  audio_enabled    NUMBER(1) DEFAULT 0 NOT NULL,       -- ★ פר-נבחן; הבוחן קובע עד האישור
  time_multiplier  NUMBER(3,2) DEFAULT 1 NOT NULL CHECK (time_multiplier IN (1, 1.25, 1.5)),
  exam_started_at  TIMESTAMP,
  examinee_token   VARCHAR2(64 CHAR),                  -- מונפק בהרשמה; חובה בכל בקשה עוקבת
  dq_count         NUMBER(2) DEFAULT 0 NOT NULL,
  has_extended_screen NUMBER(1) DEFAULT 0 NOT NULL,    -- דגל אזהרה, לא פסילה
  warning_count    NUMBER(3) DEFAULT 0 NOT NULL,       -- אזהרות שבוטלו — מוצג לבוחן
  last_warning     VARCHAR2(200 CHAR),
  finished_on_device_at TIMESTAMP                      -- "סיים — מסנכרן תוצאה" (מונע מבחן-חוזר מיותר)
);
CREATE INDEX ix_reg_session_status ON exam_registrations (session_id, status);
CREATE INDEX ix_reg_nid ON exam_registrations (national_id, registered_at);
COMMENT ON COLUMN exam_registrations.status IS 'מכונת המצבים — ראה FUNCTIONAL_SPEC §2. DISQUALIFIED חוסם הרשמה חוזרת עד הכרעת בוחן';

-- אילו שאלות נמסרו לנבחן (אימות יושרה — היום ב-CacheService, כאן טבלה אמיתית)
CREATE TABLE issued_questions (
  registration_id NUMBER NOT NULL REFERENCES exam_registrations(registration_id),
  question_id     NUMBER NOT NULL REFERENCES questions(question_id),
  position        NUMBER(2) NOT NULL,                  -- מיקום במבחן (0-29)
  shuffle_order   VARCHAR2(20 CHAR) NOT NULL,          -- פרמוטציית התשובות, למשל "2,0,3,1"
  PRIMARY KEY (registration_id, question_id)
);
COMMENT ON TABLE issued_questions IS '★ הגשה שמזכירה שאלה שלא כאן — נדחית. הניקוד: השרת גוזר נכונות מ-question_answers + הפרמוטציה';

-- הארכות זמן באמצע מבחן (ביקורת)
CREATE TABLE time_extensions_audit (
  extension_id    NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  registration_id NUMBER NOT NULL REFERENCES exam_registrations(registration_id),
  minutes_added   NUMBER(3) NOT NULL,
  reason          VARCHAR2(300 CHAR) NOT NULL,         -- חובה — הראיה בביקורת
  examiner_id     NUMBER NOT NULL REFERENCES examiners(examiner_id),
  created_at      TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL
);

-- ---------------------------------------------------------------------------
-- 4. תוצאות (לא נמחקות לעולם)
-- ---------------------------------------------------------------------------

CREATE TABLE exam_results (
  result_id        NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  registration_id  NUMBER REFERENCES exam_registrations(registration_id),
  exam_date        TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
  national_id      VARCHAR2(9 CHAR) NOT NULL,
  full_name        VARCHAR2(80 CHAR) NOT NULL,
  phone            VARCHAR2(20 CHAR),
  license_code     VARCHAR2(3 CHAR) NOT NULL,
  score            NUMBER(3),
  total_questions  NUMBER(3) DEFAULT 30 NOT NULL,
  outcome          VARCHAR2(10 CHAR) NOT NULL
                   CHECK (outcome IN ('PASS','FAIL','DISQUALIFIED','CANCELLED')),
                   -- CANCELLED = "בוטל": שורה שנדרסה ע"י הגשה/תיקון מאוחר. נשארת לתמיד
  time_taken_sec   NUMBER(6),
  session_id       NUMBER REFERENCES exam_sessions(session_id),
  examiner_id      NUMBER REFERENCES examiners(examiner_id),
  site_id          NUMBER REFERENCES sites(site_id),
  classroom        VARCHAR2(80 CHAR),
  lang_code        VARCHAR2(2 CHAR),
  lang_path        VARCHAR2(60 CHAR),                  -- מסלול החלפות שפה, למשל "he>ru>he"
  attempt_num      NUMBER(2) DEFAULT 1 NOT NULL,
  population       VARCHAR2(40 CHAR),
  audio_used       NUMBER(1) DEFAULT 0 NOT NULL,
  device_class     VARCHAR2(10 CHAR) CHECK (device_class IN ('phone','tablet','desktop')),
  verified_state   VARCHAR2(12 CHAR) DEFAULT 'VERIFIED'
                   CHECK (verified_state IN ('VERIFIED','UNVERIFIED','MANUAL')),
  is_suspicious    NUMBER(1) DEFAULT 0 NOT NULL,       -- מבחן < 3 דקות
  dq_event_id      VARCHAR2(40 CHAR),
  was_corrected    NUMBER(1) DEFAULT 0 NOT NULL,
  corrected_by     NUMBER REFERENCES examiners(examiner_id),
  correction_reason VARCHAR2(300 CHAR),
  corrected_at     TIMESTAMP,
  certificate_url  VARCHAR2(300 CHAR),
  sent_to_examinee NUMBER(1) DEFAULT 0 NOT NULL
);
CREATE INDEX ix_res_session ON exam_results (session_id);
CREATE INDEX ix_res_nid_date ON exam_results (national_id, exam_date);
CREATE INDEX ix_res_site_date ON exam_results (site_id, exam_date);
COMMENT ON COLUMN exam_results.score IS 'סף מעבר = CEIL(total_questions * 0.86) — נוסחה, לא קבוע';

-- פירוט שגויות — היום טקסט חופשי בעמודה; כאן שורות אמיתיות (שיפור)
CREATE TABLE result_wrong_answers (
  result_id       NUMBER NOT NULL REFERENCES exam_results(result_id),
  question_id     NUMBER NOT NULL,
  position        NUMBER(2),
  given_answer    VARCHAR2(1000 CHAR),
  correct_answer  VARCHAR2(1000 CHAR),
  PRIMARY KEY (result_id, question_id)
);

-- ---------------------------------------------------------------------------
-- 5. תרגול (להפריד קיבולת מהבחינות! ראה FUNCTIONAL_SPEC §11)
-- ---------------------------------------------------------------------------

CREATE TABLE classes (
  class_id    NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  class_code  VARCHAR2(10 CHAR) NOT NULL UNIQUE,
  class_name  VARCHAR2(80 CHAR) NOT NULL,
  teacher_id  NUMBER NOT NULL REFERENCES teachers(teacher_id),
  license_code VARCHAR2(3 CHAR),
  created_at  TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
  is_active   NUMBER(1) DEFAULT 1 NOT NULL,
  deleted_at  TIMESTAMP                                -- מחיקה רכה: מייתרת את גיליון "כיתות שנמחקו"
);

CREATE TABLE class_students (
  class_id    NUMBER NOT NULL REFERENCES classes(class_id),
  student_key VARCHAR2(64 CHAR) NOT NULL,              -- מזהה תלמיד (אין ת.ז. בתרגול)
  student_name VARCHAR2(80 CHAR) NOT NULL,
  joined_at   TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
  PRIMARY KEY (class_id, student_key)
);

CREATE TABLE practice_results (
  practice_id  NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  practiced_at TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
  student_key  VARCHAR2(64 CHAR) NOT NULL,
  student_name VARCHAR2(80 CHAR),
  class_id     NUMBER REFERENCES classes(class_id),    -- NULL = תרגול אנונימי
  mode         VARCHAR2(12 CHAR) NOT NULL CHECK (mode IN ('EXAM','CATEGORY','FLASHCARDS','REVIEW')),
  license_code VARCHAR2(3 CHAR),
  score        NUMBER(3),
  total        NUMBER(3),
  passed       NUMBER(1),
  time_sec     NUMBER(6),
  topic        VARCHAR2(40 CHAR),
  lang_code    VARCHAR2(2 CHAR),
  phone        VARCHAR2(20 CHAR)
);
CREATE INDEX ix_prac_class_date ON practice_results (class_id, practiced_at);
-- הגשת תרגול חייבת לאמת class_code קיים ופעיל — הפער הידוע במערכת הנוכחית (TODO 1.2)

CREATE TABLE practice_wrong_answers (
  practice_id NUMBER NOT NULL REFERENCES practice_results(practice_id),
  question_id NUMBER NOT NULL,
  category    VARCHAR2(40 CHAR),
  PRIMARY KEY (practice_id, question_id)
);

CREATE TABLE student_progress (
  student_key  VARCHAR2(64 CHAR) NOT NULL,
  class_id     NUMBER REFERENCES classes(class_id),
  streak_days  NUMBER(4) DEFAULT 0 NOT NULL,
  wrong_qs     CLOB,                                   -- JSON: שאלות שגויות מצטברות ל"חזרה על שגיאות"
  history      CLOB,                                   -- JSON: היסטוריית תרגולים
  updated_at   TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
  PRIMARY KEY (student_key)
);

CREATE TABLE risk_predictions (
  computed_at  TIMESTAMP NOT NULL,
  student_key  VARCHAR2(64 CHAR) NOT NULL,
  student_name VARCHAR2(80 CHAR),
  class_id     NUMBER REFERENCES classes(class_id),
  license_code VARCHAR2(3 CHAR),
  avg_score    NUMBER(5,2),
  practice_cnt NUMBER(5),
  trend        VARCHAR2(10 CHAR),
  pass_prob    NUMBER(4,3),                            -- הסתברות מעבר מכוילת
  risk_level   VARCHAR2(10 CHAR),
  confidence   VARCHAR2(10 CHAR),
  PRIMARY KEY (computed_at, student_key)
);
COMMENT ON TABLE risk_predictions IS 'פלט batch לילי — נבנה מחדש, לא מקור אמת';

-- =============================================================================
-- הערות מיגרציה (ראה גם REBUILD_BRIEF §4):
-- 1. ת.ז. מהגיליון עלולה להגיע כמספר בלי אפסים מובילים — לנרמל ל-9 תווים בייבוא.
-- 2. שורות היסטוריות בגיליון קצרות (עמודות נוספו עם הזמן) — קריאה הגנתית.
-- 3. אותה ת.ז. בכמה שורות ממתינים — הרשומה האחרונה קובעת; לייבא הכל, לא לדדות.
-- 4. "בוטל" בעמודת עבר/נכשל = outcome CANCELLED; אסור לדלג עליהן בייבוא.
-- 5. שני מבני "תוצאות" שונים (בחינות/תרגול) — לא לאחד בייבוא.
-- 6. אימות סיום: ספירות לפי (אתר, חודש, דרגה) חייבות להתאים לגיליון אחד-לאחד.
-- =============================================================================
