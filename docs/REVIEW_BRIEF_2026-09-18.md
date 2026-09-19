# תדריך לבדיקת מומחים — קריסות חוזרות בבוקרי בחינות, 15–17/09/2026

**נכתב:** 18/09/2026 בבוקר, על ידי הסשן שטיפל בשלושת הימים. **מטרה:** לאפשר לבדיקה עצמאית לערער על המסקנות שלי, לא לרשת אותן.

## 0. כללים לבודקים

1. **כל טענה כאן מסומנת:** `[נמדד]` — יש נתון גולמי בנספח; `[הוכח]` — נבדק בניבוי מראש או בבדיקה שהצליחה לשחזר; `[השערה]` — הסבר סביר בלי הוכחה. **תתקפו קודם את ה־`[השערה]`.**
2. הסשן הקודם (אני) תיקן בכל יום פגם אמיתי ומדוד — ובכל יום שלמחרת התקלה חזרה בצורה דומה, ~09:30. **ההנחה שאתם צריכים לבדוק: שהתיקונים היו נכונים אבל לא של הסיבה.**
3. אל תסתפקו בהסבר שמתאים לנתונים. תגידו **איזה נתון היה מפריך אותו**, ואם הוא קיים בנספח — בדקו.
4. פתרונות מחוץ למסגרת הקיימת מבוקשים במפורש (סעיף 7).

## 1. המערכת בעשר שורות

- מערכת מבחני תאוריה לצה"ל: נבחן נרשם בטלפון שלו ← בוחן מאשר בטאבלט ← 30 שאלות רצות **מקומית** במכשיר ← הגשה אחת לשרת ← ציון בשרת. הנבחן לא רואה ציון.
- לקוח: ארבעה דפי HTML סטטיים ב־GitHub Pages: `examiner.html`, `examinee.html`, `teacher.html`, `student.html` (תרגול לתלמידים).
- שרת: **סקריפט Google Apps Script אחד** (`external_exam_apps_script.js`, ~8,700 שורות, ~67 פעולות) על **גיליון Google Sheets אחד** שמכיל הכול: `ממתינים`, `תוצאות` (~4,470 שורות), `מבחנים` (JSON של מפת שאלות לכל מבחן), `בוחנים`, `סשנים`, **`תוצאות תרגול` (~107,900 שורות, ~840/יום, שתי עמודות JSON גדולות)**, `אבחון` (יומן).
- מגבלות הפלטפורמה: **30 הרצות בו־זמנית לחשבון**, **הריגה ב־360 שניות**, CacheService 100KB לערך ו־6 שעות TTL לא מובטחות, חשבון gmail.com רגיל (90 דק' טריגרים ביום).
- סקרים מהמכשירים (כולם משורשרים — בקשה אחת באוויר לכל סקר): נבחן ממתין `checkApproval` כל 5 שנ' (נסיגה עד 20); נבחן במבחן `getExamStatus` כל 10 שנ' (עד 20); בוחן `examinerDashboard` כל 5 שנ' (2 שנ' בזמן סנכרון תוצאה).
- דיאגנוסטיקה בשרת: גיליון `אבחון` רושם בקשה שנמשכה **≥15 שניות** (`SLOW`) עם סימוני שלב `phase@ms`, והרצות שנהרגו (`KILLED`). **בקשות מהירות לא נרשמות — נפח העומס לא נראה שם.**
- פריסה: לקוח = `git push` (Pages); שרת = הדבקה ידנית + "גרסה חדשה". **כל דחיפה ל־Pages מחתימה מחדש ETag לכל הקבצים.**
- מקביל: המערכת נכתבת מחדש ב־Oracle APEX (צוות ה־APEX) — לוח זמנים לא ידוע לי.

## 2. ציר הזמן (שעון ישראל) — מה קרה ומה נמדד

### 15/09 — יום בחינות + בדיקות של המפעיל אחה"צ
- בוקר: בוחן במחנה עמוס — "בסה"כ תקין", שני אייפונים לא עברו מממתינים למבחן. `[דיווח]`
- 12:00–12:30: `health` (פעולה שלא עושה כלום) על הסקריפט שלנו: **100.8s / 69.3s / 30.4s → 404** דף שגיאה של גוגל; 5 מקבילים: 4×404 + 1×200, כולם ~30s. **שלושה סקריפטים אחרים באותו חשבון באותה דקה: 1.7s / 5.8s.** דף "ביצועים": הרצה **אחת** בלבד במצב "פועלת" (רוח רפאים מ־12/09, 72 שעות). `[נמדד]`
- ניבוי מראש שנכשל: אם הזמן בתוך הקוד שלנו, יופיעו שורות `SLOW GET health` — **לא הופיעו** (בדיקת health ב־12:25 לקחה 36.5s בדפדפן/curl ולא נרשמה) → הזמן **מחוץ להרצה**, בשכבת ההגשה של גוגל לדפלוי הזה. `[הוכח]` תעתיק ההסבר: `googleusercontent.com/macros/echo` הוא שמחזיר 404 אחרי 30–100s; ה־302 הראשון מהיר.
- אחה"צ: המפעיל פותח לוח מפקד ~15 פעמים ב־31 דקות (31–56s כל פתיחה); `examinerDashboard` באותו חלון: **78.6s, 54.4s**. `[נמדד]` מבחן בדיקה של המפעיל: הגשה 35s, `drive:he` נעלם (תיקון r15 הוכח), 32s אחרי הסימון האחרון (חמש קריאות גיליון מלאות). `[נמדד]`

### 16/09 — יום בחינות. **קריסה ארצית ~09:30, מעבר לנייר**
- 07:26 ו־07:57: שתי דחיפות לקוח (ביטול טעינה אוטומטית של לוח מפקד; תיקון ארגומנט timeout שנפל בשקט). **אף דחיפה אחרי 07:57** (ה־API של GitHub מאשר). `[הוכח]`
- ~09:30: כל הבוחנים, **באזורים שונים ברחבי הארץ**, נזרקו למסך כניסה עם "שגיאת תקשורת. נסה שוב". ~20 דק' אח"כ חלקית; דפים שחזרו **רעננו את עצמם בלי הפסקה** תחת "גרסה חדשה זמינה — מתעדכן אוטומטית". דף מורה, דף תלמיד, ונבחנים שכבר בתוך מבחן — עבדו; תוצאות הגיעו לשרת, הבוחן לא ראה אותן. בוחן: "הכול התחיל מעדכון אוטומטי באמצע בחינה". **אף אחד לא ראה "פג תוקף ההתחברות".** `[דיווח]`
- נמדד באותו יום: **אותו build מוגש כ־`"6aaa21d2-76fe0"` (ללא דחיסה) או `W/"6aaa21d2-76fe0"` (gzip)**; בדיקת העדכון העצמי השוותה אותם כמחרוזת, ספרה גם דפי שגיאה כ"גרסה", ורעננה באמצע סשן. `[נמדד]` הסקריפט של הבדיקה מיירט גם HEAD (ההערה בקוד טענה אחרת).
- ההסבר שנכתב: שרת ההפצה (Fastly) התחיל להחזיר את הכתיב השני **לכל הארץ** ← כל דף בוחן ראה "גרסה חדשה" באותן 2 דקות ← רענון מסונכרן ← מצב אבד ← כניסה נכשלת תחת עומס. `[השערה]` — **מתאימה לארבע העובדות (ארצי, מסונכרן, בלי פריסה, הבאנר ראשון) אבל לא נצפתה ישירות; מדידות מאוחרות הראו כתיב יציב.**
- `אבחון` 16/09: **אפס שורות 09:00–11:19** (חלון התקלה!) — דו־משמעי: או שבקשות לא הגיעו לקוד, או שגיליון האבחון עצמו לא היה ניתן לכתיבה. מ־11:19: קריאות של גיליונות קטנים לוקחות **22s, 29s, 65s**; `loadStudentProgress` (תרגול) **100s**; `studentJoinClass` 35s. `[נמדד]`
- **ממצא נוסף מאותו יום:** מאז 01/06 כל רענון של לוח הבוחן (5s) הריץ את `siteCombinedReport` **המלא** (3 קריאות גיליון) רק כדי להחליט אם להציג כפתור — ~1,000 הרצות לבוחן לבוקר. `[הוכח בקוד ובבדיקה]`

### 17/09 — יום בחינות. **תקיעה ~09:45, בלי רענון־לולאה**
- ~09:15 שני בוחנים באזורים שונים מתחילים; ~09:45 שניהם רואים **"שגיאת תקשורת" ורעננו בעצמם**; אחרי זמן־מה חזרו לעבוד. נבחני אייפון לא הצליחו להתחבר בספארי, עברו לכרום (מה הופיע במסך — לא ידוע). אותם טאבלטים. אף דחיפה. `[דיווח]`
- `אבחון` 17/09 (נספח ב'): `examinerDashboard` **88.7s — 84.8 לפני הסימון הראשון**; `examinerDashboard` **96.7s — כל הקריאות נגמרו ב־3.8s ואז 93s של כלום**; `registerExamQuestions` 121s; **`getExamStatus` 354s — נהרג** (המטפל: קריאת זנב 1000 שורות + גיליון קטן); `listActiveExaminers` 196s (קריאה אחת קטנה); 12 הגשות 15–18s, כולן הגיעו; `teacherClassDetails` **13 פעמים** בחלון, 19–32s כל אחת (קריאה מלאה של 107k שורות, בלחיצת מורה). `[נמדד]`

### 18/09 — r21 נפרס: המטמון מחמם את עצמו (בדיקה שעתית, אפס עבודה ידנית).

## 3. מה תוקן — ומה כל תיקון **הוכח** לעשות

| commit | תיקון | מה הוכח |
|---|---|---|
| `9b24cd3` r15 | הגשת תוצאה בלי Drive; סימון חינמי <8s; מסך "אין סשן" לא מסתיר כשל טעינה | `drive:he` נעלם במבחן חי |
| `6f6199f` r19 | לוח מפקד: קריאה חסומה־בתאריך של תוצאות/תרגול (107k→26.8k שורות) | `rows26775/107637` בטרייל; ~2× מהירות |
| `9cab4b9` | לוח מפקד לא נטען אוטומטית בכניסה של בוחן־מפקד | קוד |
| `68cfcb4` | `apiGet` בדף הבוחן לא העביר timeout — לוח מפקד רץ ב־30s במקום 90 | דוחות נפתחו מיד אחרי |
| `49a9688` | בדיקת העדכון העצמי בארבעת הדפים: ETag מנורמל, רק 200, שני סקרים רצופים, **אין רענון עם סשן פתוח** | הבדיקה משחזרת את 16/09 על הקוד הישן; ב־17/09 לא הייתה לולאה |
| `c780e05` | הדו"ח המשותף לא רץ בכל רענון (פעם ב־10 דק') | קוד + בדיקה (60 רענונים → 1 דו"ח) |
| `b370eb0` r21 | מטמון שומר על עצמו; סימונים ב־getExamStatus/teacherClassDetails | יומן התקנה תקין |

**מה שלא תוקן ורשום כפתוח:** באגים בכניסת בוחן (מגבלת 5 אסימונים / דריסה / מחיקת "זכור אותי" בשגיאה רגעית); `sumExtraMinutes` קורא גיליון שלם בכל סקר סטטוס; שישה מטפלים קוראים את 107k שורות התרגול במלואן; הגשה = שלוש קריאות מלאות של `תוצאות` (~13s); לוח מפקד עדיין ~30s.

## 4. הדפוס שלא הוסבר

| | 15/09 | 16/09 | 17/09 |
|---|---|---|---|
| שעה | 12:00–12:30 (אחה"צ, בדיקות) | ~09:30 | ~09:45 |
| סימפטום בשטח | health 30–100s/404 | מסך כניסה + לולאת רענון | "שגיאת תקשורת" |
| סקריפטים אחרים בחשבון | **מהירים** | לא נבדק | לא נבדק |
| הרצות "פועלת" | 1 (רוח רפאים) | לא נבדק | **לא נבדק** |
| עבודה בתוך המטפלים | אפס (health) | לא נראה (אין שורות) | זניחה, אבל ההרצה הורעבה דקות |
| תעבורת תרגול/מורים | ? | כן (100s, 35s) | כן (13× 107k) |

**המשותף: הרצות שעושות כלום או כמעט כלום נתקעות לדקות, בעיקר סביב 09:30 ביום בחינות.**

## 5. ההשערות על השולחן — בעד ונגד

**א. גוגל (הידרדרות אזורית / הגשה).** בעד: 15/09 הוכח שהזמן מחוץ להרצה; 14/09 (תיעוד קודם) אותו דבר עם צומת ישראל שנפל. נגד: ב־15/09 סקריפטים אחרים באותו חשבון היו מהירים — זה **ספציפי לדפלוי הזה**, מה שלא מסתדר עם "גוגל כללי". `[השערה]`

**ב. חריגה מ־30 הרצות בו־זמנית → תור.** כל מכשיר מחזיק בקשה אחת באוויר; ברגע שהשהיית השרת עוברת את מרווח הסקר, מספר ההרצות הפעילות ≈ מספר המכשירים (40+ בשני סשנים) > 30 → תור → הזמן בתור נראה כ"85 שניות לפני הסימון הראשון" → לולאה שמזינה את עצמה. **מסביר 17/09 09:18 בדיוק.** לא מסביר את 09:22 (93s **בתוך** ההרצה, אחרי הקריאות — הרעבת CPU?). `[השערה]` **המפריך: דף "ביצועים" 17/09 09:15–09:40 — כמה הרצות חופפות.**

**ג. תחרות על הגיליון המשותף.** קריאות מלאות של 107k שורות (מורים, תלמידים) על אותו גיליון של הבחינות; ב־16/09 וב־17/09 היו כאלה בחלון. ב־16/09 קריאת גיליון של 5 שורות לקחה 65s. `[נמדד]` שזה **תואם**, `[השערה]` שזה **הסיבה**. המפריך: אותו דף "ביצועים" — אם אין תרגול/מורים ב־09:15–09:45, נפל.

**ד. משהו שלא ראיתי.** למשל: `diagFinish` כותב לגיליון האבחון בתוך נתיב התשובה של כל בקשה ≥15s (נפסל ב־15/09 בניבוי מראש — אבל רק לסוג בקשה אחד); ה־SW של הבוחן; מגבלת קצב; הרוח־רפאים שרצה 72 שעות.

## 6. נתונים חסרים שיכריעו

1. **דף "ביצועים" בעורך, 17/09 09:15–09:40** (וגם 16/09 09:20–10:00): כמה הרצות, אילו פונקציות, כמה זמן, כמה חופפות. **זה הנתון היחיד שמפריד בין ב', ג' ו־א'.**
2. **מה הופיע במסך בספארי** אצל נבחני האייפון (מסך לבן / "נסה שוב" / שגיאה).
3. האם ב־16/09 בין 09:00 ל־11:19 גיליון האבחון לא נכתב כי לא היו בקשות איטיות, או כי הכתיבה נכשלה (`diagFinish` בולע שגיאות).

## 7. שאלות מחוץ למסגרת — מה שאני מבקש שתדונו בו

1. **הפרדת התרגול (תלמידים/מורים) לגיליון ולסקריפט משלו — עכשיו,** לא בשכתוב. זה מוריד 107k שורות ואת כל הקריאות המלאות מהגיליון של הבחינות ומחלק את 30 ההרצות.
2. **תכנון הסקרים:** 40 מכשירים × סקר כל 5–10 שניות מול תקרה של 30 הרצות. האם הפתרון הוא נסיגה אגרסיבית יותר (כבר יש ×1.5 עד 20s), סקר מאוחד, או שינוי ארכיטקטוני (למשל מצב סשן ב־CacheService במקום בגיליון)?
3. **האם להאיץ את המעבר ל־APEX** במקום להמשיך לחזק — ואם כן, מה המינימום שצריך להחזיק עד אז.
4. האם יש דרך **לראות בזמן אמת** כמה הרצות רצות (למשל מונה ב־CacheService ב־doGet/doPost), כדי שהפעם הבאה לא תדרוש צילום מסך בדיעבד.

## 8. איפה הדברים

- ריפו: `standalone_exam/` (ריפו git משלו, `master`, **ציבורי**). HEAD = `b370eb0`. שרת חי = `2026-09-17-r21` (`health&origin=examinee-app`). לקוח חי על Pages = `c780e05` בדפי ה־HTML (הדחיפות אחרי הן שרת+בדיקות בלבד).
- בדיקות: `standalone_exam/tests/*.test.cjs` — `node tests/hot_path_hardening.test.cjs` (79), `client_reliability` (61), `api_reliability` (8), `warmup_budget` (20), `database_reliability` (35), `cache_reliability`. שלוש חבילות ישנות (`cache_lease_reliability`, `result_lease_cleanup`, `submission_commit`) נכשלות מאז לפני 15/09 — לא רגרסיה.
- טריגרים כרגע: `ensureQuestionCachesWarm` כל שעה; `archiveOldPendingRows` יומי 01:00; `rebuildAtRiskCache` יומי 03:00.
- דיאגנוסטיקה: גיליון `אבחון`; מדידת תשתית: `curl -sL -w '%{http_code} %{time_total}' "<exec>?action=health&origin=examinee-app"` (בלי `origin` הבקשה נדחית בשער ולא מגיעה ל־health — זה מבחן תשתית בלבד).

---

## נספח א' — `אבחון` 16/09 (UTC; ישראל = +3). אפס שורות בין 06:00 ל־08:19.

```
2026-09-16T08:19:39.986Z  SLOW GET  loadStudentProgress   30052
2026-09-16T08:20:26.054Z  SLOW GET  studentJoinClass      35486
2026-09-16T08:28:12.275Z  SLOW GET  loadStudentProgress  100020
2026-09-16T08:30:22.435Z  SLOW GET  teacherClassDetails   15966
2026-09-16T08:30:49.898Z  SLOW GET  studentJoinClass      36063
2026-09-16T08:34:05.303Z  SLOW GET  examinerDashboard     24655  sheet:pending-dash@384 sheet:results-dash@22458 sheet:extensions-dash@23307 sheet:results-dash-2@23605 compute:dash-done@24547
2026-09-16T08:37:49.889Z  SLOW GET  getExamQuestions      15099
2026-09-16T09:01:59.121Z  SLOW GET  examinerDashboard     32446  sheet:pending-dash@491 sheet:results-dash@931 sheet:extensions-dash@30120 sheet:results-dash-2@31684 compute:dash-done@32375
2026-09-16T09:13:24.128Z  SLOW GET  teacherClassDetails   18826
2026-09-16T09:13:55.999Z  SLOW GET  teacherClassDetails   22714
2026-09-16T09:15:37.176Z  SLOW GET  teacherClassDetails   21601
2026-09-16T09:44:40.325Z  SLOW GET  siteCombinedReport    65948  sheet:sessions-report@65372 sheet:role-report@65645
2026-09-16T09:47:18.740Z  SLOW GET  siteCombinedReport    18100  sheet:sessions-report@17597 sheet:role-report@17917
2026-09-16T09:49:03.438Z  SLOW GET  studentJoinClass      29749
```
(`siteCombinedReport` כאן = הבדיקה האוטומטית מלוח הבוחן, לא פתיחה ידנית — תוקן ב־`c780e05`.)

## נספח ב' — `אבחון` 17/09 (UTC; ישראל = +3)

```
2026-09-17T05:56:20.678Z  SLOW GET  teacherClassDetails    23981
2026-09-17T06:07:01.759Z  SLOW GET  teacherClassDetails    24836
2026-09-17T06:18:31.700Z  SLOW GET  examinerDashboard      88715  sheet:pending-dash@84837 sheet:results-dash@85899 sheet:extensions-dash@87426 sheet:results-dash-2@87748 compute:dash-done@88650
2026-09-17T06:22:13.039Z  SLOW GET  examinerDashboard      96665  sheet:pending-dash@558 sheet:results-dash@2550 sheet:extensions-dash@3525 sheet:results-dash-2@3824 compute:dash-done@96588
2026-09-17T06:22:29.049Z  SLOW GET  teacherClassDetails    30277
2026-09-17T06:24:52.254Z  SLOW GET  getExamStatus          62289
2026-09-17T06:24:52.301Z  SLOW POST registerExamQuestions 120990
2026-09-17T06:25:27.677Z  SLOW POST registerExamQuestions  46489
2026-09-17T06:30:28.934Z  SLOW GET  getExamStatus         354453
2026-09-17T06:34:38.316Z  SLOW POST submitResult          16691  sheet:token-submit@169 sheet:pending-submit@682 sheet:registered-submit@818 meta:wrong-answers@2225 sheet:results-submit@3293 sheet:results-submit-3@11462 compute:submit-done@16624
2026-09-17T06:37:09.681Z  SLOW POST submitResult          15668  sheet:token-submit@167 sheet:pending-submit@2440 sheet:registered-submit@2711 meta:wrong-answers@3611 sheet:results-submit@4858 sheet:results-submit-3@11370 compute:submit-done@15523
2026-09-17T06:56:34.441Z  SLOW POST submitResult          33151  sheet:token-submit@80
2026-09-17T07:05:57.254Z  SLOW POST submitResult          17649  sheet:token-submit@60 sheet:pending-submit@876 sheet:registered-submit@1076 meta:wrong-answers@1999 sheet:results-submit@3692 sheet:results-submit-3@8624 compute:submit-done@17562
2026-09-17T07:08:07.593Z  SLOW GET  getExamStatus          58270
2026-09-17T07:09:47.796Z  SLOW GET  teacherClassDetails    28080
2026-09-17T07:09:58.433Z  SLOW GET  teacherClassDetails    25276
2026-09-17T07:11:13.053Z  SLOW GET  teacherClassDetails    24519
2026-09-17T07:16:30.802Z  SLOW GET  getExamStatus          88235
2026-09-17T07:16:40.992Z  SLOW POST registerExamQuestions  18675
2026-09-17T07:20:48.374Z  SLOW POST submitResult          18304  sheet:token-submit@66 sheet:pending-submit@1126 sheet:registered-submit@1312 meta:wrong-answers@2182 sheet:results-submit@3057 sheet:results-submit-3@6934 compute:submit-done@18237
2026-09-17T07:22:50.400Z  SLOW POST submitResult          16716  sheet:token-submit@56 sheet:pending-submit@906 sheet:registered-submit@1104 meta:wrong-answers@2186 sheet:results-submit@3278 sheet:results-submit-3@12138 compute:submit-done@16652
2026-09-17T07:24:09.971Z  SLOW POST submitResult          17069  sheet:token-submit@87 sheet:pending-submit@1156 sheet:registered-submit@1400 meta:wrong-answers@2760 sheet:results-submit@3631 sheet:results-submit-3@11955 compute:submit-done@17008
2026-09-17T07:24:37.828Z  SLOW GET  teacherClassDetails    30363
2026-09-17T07:26:14.353Z  SLOW GET  teacherClassDetails    19019
2026-09-17T07:26:21.053Z  SLOW GET  teacherClassDetails    25832
2026-09-17T07:26:32.320Z  SLOW GET  teacherClassDetails    25310
2026-09-17T07:26:43.412Z  SLOW GET  teacherClassDetails    25261
2026-09-17T07:26:52.729Z  SLOW GET  teacherClassDetails    32041
2026-09-17T07:26:57.804Z  SLOW GET  teacherClassDetails    18789
2026-09-17T07:28:50.916Z  SLOW GET  teacherClassDetails    32277
2026-09-17T07:30:03.125Z  SLOW GET  examinerDashboard     160095  sheet:pending-dash@936 sheet:results-dash@64961 sheet:extensions-dash@158117 sheet:results-dash-2@158499 compute:dash-done@159987
2026-09-17T07:37:50.574Z  SLOW POST submitResult          15233  sheet:token-submit@51 sheet:pending-submit@864 sheet:registered-submit@1126 meta:wrong-answers@2056 sheet:results-submit@2927 sheet:results-submit-3@11847 compute:submit-done@15187
2026-09-17T07:42:31.875Z  SLOW GET  examinerDashboard      95843  sheet:pending-dash@92251 sheet:results-dash@93340 sheet:extensions-dash@94345 sheet:results-dash-2@94773 compute:dash-done@95786
2026-09-17T07:44:57.751Z  SLOW POST submitResult          15788  sheet:token-submit@63 sheet:pending-submit@894 sheet:registered-submit@1584 meta:wrong-answers@2826 sheet:results-submit@3976 sheet:results-submit-3@11901
2026-09-17T07:45:20.698Z  SLOW GET  getExamStatus          92915
2026-09-17T07:48:57.627Z  SLOW GET  getExamStatus          93508
2026-09-17T07:53:43.904Z  SLOW POST submitResult          17018  sheet:token-submit@61 sheet:pending-submit@1000 sheet:registered-submit@1285 meta:wrong-answers@3109 sheet:results-submit@4334 sheet:results-submit-3@12611 compute:submit-done@16971
2026-09-17T07:55:24.705Z  SLOW GET  listActiveExaminers   196106
2026-09-17T07:58:32.752Z  SLOW GET  teacherClassDetails    25590
2026-09-17T08:01:03.390Z  SLOW POST submitResult          18308  sheet:token-submit@86 sheet:pending-submit@1026 sheet:registered-submit@1308 meta:wrong-answers@2697 sheet:results-submit@4439 sheet:results-submit-3@13646 compute:submit-done@18257
```
הערה לקריאה: הזמן ב־`@ms` הוא מתחילת `doGet`/`doPost` — **זמן בתור לפני תחילת ההרצה לא נמדד.** ב־06:18 84.8s לפני `sheet:pending-dash` = בתוך ההרצה, לפני הקריאה הראשונה (`requireToken` קורא את גיליון `בוחנים` במלואו; `getSheet` פותח את הגיליון).

## נספח ג' — מדידות תשתית (curl מהמחשב של המפעיל)
- 15/09 12:00–12:30: exec שלנו 30–100s/404; סקריפטים אחרים באותו חשבון 1.7s/5.8s; 12:17 → 1.9/10.6/9.2s; 13:00 → 2.1s.
- 16/09 07:09: 1.8–3.0s. 07:42: 1 מתוך 9 נתקע 60s, השאר 1.8–2.9s. 16/09 ETag ל־examiner.html: 8/8 `W/"6aaa21d2-76fe0"` בבקשה דמוית־דפדפן, `"6aaa21d2-76fe0"` ב־curl רגיל.
- 17/09 07:09: 1.8–2.6s. 18/09 05:39: 1.9–2.8s.
