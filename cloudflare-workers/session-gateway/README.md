# session-gateway — שער סקרים מאחד

Worker שמאחד את סקרי הנבחנים: **בקשה אחת לסשן כל 3 שניות** ל‑Apps Script, כמה נבחנים שלא יסקרו.
במקום ~480 הרצות לדקה ל‑40 נבחנים ממתינים — 20 לדקה לסשן. (DESIGN_2026-09-21 §3.4)

## נקודות קצה

| נתיב | תשובה |
|---|---|
| `GET /` | `{status:'ok', service:'session-gateway', build:'…'}` |
| `GET /v1/poll?kind=approval&sessionCode=ABC12345&idNumber=…&examineeToken=…` | בדיוק כמו `checkApproval` בשרת |
| `GET /v1/poll?kind=status&…` | בדיוק כמו `getExamStatus` בשרת |

התשובות זהות לאלו של השרת — הלקוח מחליף כתובת בלבד, לא לוגיקה. שדה נוסף אחד: `stale:true`
כשהתשובה חושבה מעותק ישן (עד 60 שניות) כי ה‑upstream נכשל. כשאין שום עותק:
`{status:'error', code:'upstream_unavailable', retryable:true}` ב‑HTTP 200 — הלקוח מאט,
**לא** נופל חזרה לסקר ישיר (אחרת סערת הבקשות חוזרת).

הטוקן של הנבחן לא עוזב את Apps Script: ה‑snapshot מחזיק `tokenHash` (SHA‑256), וה‑Worker
מגבב את מה שהלקוח שלח ומשווה.

## פריסה

```bash
cd cloudflare-workers/session-gateway
npx wrangler secret put GATEWAY_KEY     # אותו ערך בדיוק כמו ה-ScriptProperty
npx wrangler deploy
```

פריסה מה‑dashboard מחזירה 403 בחשבון הזה — wrangler בלבד (כמו `exam-results-worker`).

## מה צריך בצד השרת (ScriptProperties)

| Property | ערך |
|---|---|
| `GATEWAY_KEY` | הסוד שמאפשר ל‑Worker לקרוא `action=sessionSnapshot` (אותו ערך כמו ב‑wrangler secret) |
| `GATEWAY_URL` | `https://session-gateway.<account>.workers.dev` — `getSessionInfo` מחזיר אותו ללקוח. **ריק = כל הלקוחות חוזרים לסקר ישיר**, וזה מתג הכיבוי בלי דחיפת Pages |

`API_URL` ב‑`wrangler.toml` חייב להיות כתובת ה‑`/exec` הפעילה של הסקריפט.

## אימות אחרי פריסה

```bash
curl -s https://session-gateway.<account>.workers.dev/
# {"status":"ok","service":"session-gateway","build":"2026-09-21"}

curl -s "https://session-gateway.<account>.workers.dev/v1/poll?kind=approval&sessionCode=<קוד>&idNumber=<ת.ז.>"
# {"status":"error","message":"לא נמצא רישום"} לפני הרשמה, ואחריה {"status":"ok","approval":"waiting",…}
```

שתי בקשות ברצף לאותו סשן צריכות להחזיר את אותה תשובה בלי הרצת Apps Script שנייה — נבדק
בגיליון ההרצות: ‏`sessionSnapshot` אחד לכל ~3 שניות של סקר, לא אחד לכל נבחן.

## תקציב

100k בקשות/יום **לכל החשבון** — משותף עם ה‑TTS, ה‑proxy לתמונות ו‑worker הדו"חות.
בוקר בחינות של 40 נבחנים ≈ 18k בקשות; שלושה אתרים ≈ 55k. מעל התקציב Cloudflare מחזיר 5xx,
והלקוח נופל חזרה לסקר ישיר אחרי 3 כשלים רצופים.

## בדיקות

`node tests/gateway.test.cjs` (בלי רשת: fetch מזויף וסָפוּר, שעון מזויף).
`package.json` כאן קיים רק כדי ש‑`worker.js` ייטען כ‑ESM ב‑Node.
