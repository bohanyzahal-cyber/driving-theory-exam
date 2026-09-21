# session-gateway — שער סקרים מאחד + מאגר השאלות הפרטי

Worker עם שני תפקידים (DESIGN_2026-09-21 §3.4 ו‑§11):

1. **מאחד את סקרי הנבחנים:** **בקשה אחת לסשן כל 2 שניות** ל‑Apps Script, כמה נבחנים שלא יסקרו.
   במקום ~480 הרצות לדקה ל‑40 נבחנים ממתינים — 30 לדקה לסשן.
2. **מגיש את טקסטי השאלות.** הבנק **אינו ציבורי**: הוא נכס סטטי פרטי של ה‑Worker
   (`run_worker_first`), וכל מכשיר מקבל **רק את המזהים שהאישור החתום שלו נוקב בהם**.

## נקודות קצה

| נתיב | מי מורשה | תשובה |
|---|---|---|
| `GET /` | — | `{status:'ok', service:'session-gateway', build:'…', bank:'<manifest.build>'}` |
| `GET /v1/poll?kind=approval&sessionCode=ABC12345&idNumber=…&examineeToken=…` | — | בדיוק כמו `checkApproval` בשרת |
| `GET /v1/poll?kind=status&…` | — | בדיוק כמו `getExamStatus` בשרת |
| `GET /v1/bank?grant=…` | scope `exam`/`practice` | המזהים שב‑grant, בכל 7 השפות |
| `GET /v1/bank?grant=…&ids=1,2,3[&langs=he,en]` | scope `examiner` | עד 60 מזהים; `langs` מסנן שפות |
| `GET /v1/bank/full?grant=…&lang=he` | scope `examiner` בלבד | הבנק המלא של השפה, כזרם |
| `POST /v1/invalidate?sessionCode=ABC12345` | — | `{"status":"ok"}` — מוחק את ה‑snapshot של הסשן |
| `POST /v1/invalidate?sessionCode=…&idNumber=…&status=approved` <br>`[&examMinutes=50&extraMinutes=0&audio=on]` | — | `{"status":"ok","patched":true}` — כותב את ההחלטה **לתוך** ה‑snapshot |

תשובת הבנק: `{"status":"ok","build":"…","questions":[…],"missing":[…]}` — `questions` הוא
תוכן `assets/q/<id>.json` **כמות שהוא** (ללא `JSON.parse` בנתיב החם), לפי סדר המזהים;
`missing` הם מזהים בלי קובץ. `Cache-Control: no-store` בכל התשובות.

תשובות הסקר זהות לאלו של השרת — הלקוח מחליף כתובת בלבד, לא לוגיקה. שדה נוסף אחד: `stale:true`
כשהתשובה חושבה מעותק ישן (עד 60 שניות) כי ה‑upstream נכשל. כשאין שום עותק:
`{status:'error', code:'upstream_unavailable', retryable:true}` ב‑HTTP 200 — הלקוח מאט,
**לא** נופל חזרה לסקר ישיר (אחרת סערת הבקשות חוזרת).

הטוקן של הנבחן לא עוזב את Apps Script: ה‑snapshot מחזיק `tokenHash` (SHA‑256), וה‑Worker
מגבב את מה שהלקוח שלח ומשווה.

### `/v1/invalidate` — דחיפה אחרי החלטת בוחן

אחרי אישור/דחייה/איפוס/פסילה/הארכת זמן, דף הבוחן שולח `POST` (fire‑and‑forget, `keepalive`).
שתי צורות:

**מחיקה (`?sessionCode=X` בלבד).** ה‑Worker מוחק את ה‑snapshot של הסשן מהזיכרון ומשתי רשומות
ה‑Cache (`snap`, `stale`), כך שהסקר הבא של הנבחן קורא מהשרת מיד במקום לחכות לתום חלון ה‑2
שניות. **מוגבל לקריאה חוזרת כפויה אחת לכל 2 שניות לסשן** — אותו תקציב שממנו שואב גם הרענון
האוטומטי של "שורה שחסרה ב‑snapshot", כך שסופת `invalidate` לא יכולה להכפיל את העומס על
Apps Script. התשובה תמיד `{"status":"ok"}`, גם כשהמגבלה בלמה את המחיקה.

**טלאי (`&idNumber=…&status=…`).** הנדנוד נושא את ההחלטה עצמה. ה‑Worker לוקח את ה‑snapshot
הנוכחי של הסשן (זיכרון → `snap` → `stale`), מוצא את השורה **האחרונה** של אותה ת.ז. (סדר
הגיליון הוא הישן ראשון), כותב בה `status` — ואם נמסרו ותקינים גם `examMinutes`, `extraMinutes`
ו‑`audio` — ושומר את ה‑snapshot כ**טרי** (זיכרון + `snap` + `stale`, אותו גוף). הסקר הבא של
הנבחן, **כשנייה אחרי הלחיצה**, עונה את ההחלטה **בלי שום קריאה לשרת**. התשובה:
`{"status":"ok","patched":true}`.

*למה זה לא הקדמת האמת:* דף הבוחן יורה את הנדנוד **רק אחרי** ש‑Apps Script ענה `status:'ok'`,
כלומר אחרי שהשורה כבר נכתבה בגיליון — הטלאי אינו יכול להקדים את המצב האמיתי ביותר מאותה
כתיבה מאושרת. הוא גם קצר‑חיים: העותק פג על שעון ה‑`FRESH_MS` הרגיל, והסקר שאחריו קורא את
השורה מהשרת וממילא דורס אותו. זה מה שהופך את "2–4 שניות" של DESIGN §11.8 ל‑**כשנייה**
באישור, בלי ערוץ push; הרצפה שנשארת היא הכתיבה של Google עצמה (~1 שנ').

טלאי הוא **כתיבה, לא קריאה**, ולכן אינו צורך את תקציב הקריאה החוזרת — `invalidate` רגיל מיד
אחריו עדיין מוחק. כשאין snapshot לסשן, או שאין בו שורה לאותה ת.ז. (הנבחן נרשם אחרי הצילום) —
נפילה חזרה למחיקה ותשובה `{"status":"ok","patched":false}`; הסקר של הנבחן ממילא קורא מהשרת
כשהשורה חסרה, ומוצא אותה בעצמו.

| פרמטר | חובה | ערכים |
|---|---|---|
| `sessionCode` | כן | `^[A-Z0-9]{6,8}$` — אחרת **400** |
| `idNumber` + `status` | רק לטלאי, ושניהם יחד | `status` אחד מתוך `waiting, approved, rejected, cancelled, in_exam, completed, disqualified, dq_confirmed`; `idNumber` חייב להכיל ספרות. ערך אחר → **400**. רק אחד מהשניים → מחיקה רגילה עם `patched:false` |
| `examMinutes` | לא | שלם 1–600; ערך לא תקין מתעלמים ממנו (שאר הטלאי עדיין נכתב) |
| `extraMinutes` | לא | שלם 0–600, כנ"ל |
| `audio` | לא | `on` / `off` בדיוק, כנ"ל |

### האישור החתום (grant)

מחרוזת `<payload>.<sig>`: `payload` = base64url של JSON, `sig` = base64url של
HMAC‑SHA256(`GATEWAY_KEY`, `payload`). **רק Apps Script חותם** (`startExam`, `startPractice`,
`bankGrant`); ה‑Worker מאמת ב‑`crypto.subtle.verify` (השוואה בזמן קבוע) ולעולם לא סומך על
הלקוח. ה‑payload: `{"v":1,"s":"exam","ids":[14,120],"sub":"ABC12345:012345678","exp":<ms>}`.

grant לא תקין / פג תוקף / scope לא מתאים → **HTTP 403** `{"status":"error","code":"grant_invalid"}`.
אין נכסים (binding חסר או כל הקריאות נכשלו) → **HTTP 503**
`{"status":"error","code":"bank_unavailable","retryable":true}`.

## הנכסים (`assets/`)

```
assets/manifest.json      {build, generatedAt, questions, langs:{<lang>:{sha,count,bytes}}}
assets/q/<id>.json        שאלה אחת, כל השפות: {"id":14,"l":{"he":{t,a,i,v?},"ru":{…},…}}
assets/bank/<lang>.json   הבנק המלא של שפה: [{id,t,a,i,v?},…] ממוין לפי id
```

**`node tools/build_bank.js` חייב לרוץ לפני כל `npx wrangler deploy`** — הוא מה שבונה את
`assets/`. התיקייה **ב‑`.gitignore`**: הטקסטים לא בריפו ולא ב‑Pages, ולכן קלון טרי (בלי
`deployment/generated/questions_<lang>.json`) **אינו יכול** לבנות אותה, ופריסה ממנו תעלה בנק ריק.

## פריסה

```bash
node tools/build_bank.js                # בונה את assets/ (מהתיקייה הראשית של הריפו)
cd cloudflare-workers/session-gateway
npx wrangler secret put GATEWAY_KEY     # אותו ערך בדיוק כמו ה-ScriptProperty
npx wrangler deploy                     # מעלה את worker.js ואת כל תיקיית assets/
```

פריסה מה‑dashboard מחזירה 403 בחשבון הזה — wrangler בלבד (כמו `exam-results-worker`).

## מה צריך בצד השרת (ScriptProperties)

| Property | ערך |
|---|---|
| `GATEWAY_KEY` | הסוד שמאפשר ל‑Worker לקרוא `action=sessionSnapshot` **וגם** המפתח שבו נחתמים ה‑grants (אותו ערך כמו ב‑wrangler secret) |
| `GATEWAY_URL` | `https://session-gateway.<account>.workers.dev` — `getSessionInfo` מחזיר אותו ללקוח. **ריק = כל הלקוחות חוזרים לסקר ישיר**, וזה מתג הכיבוי בלי דחיפת Pages |

`API_URL` ב‑`wrangler.jsonc` חייב להיות כתובת ה‑`/exec` הפעילה של הסקריפט.

## אימות אחרי פריסה

```bash
curl -s https://session-gateway.<account>.workers.dev/
# {"status":"ok","service":"session-gateway","build":"2026-09-21","bank":"3dbb0d70…"}
#  ^ אם "bank" ריק — הנכסים לא עלו: הרץ build_bank.js ופרוס שוב

curl -s -o /dev/null -w '%{http_code}\n' "https://session-gateway.<account>.workers.dev/v1/bank?grant=bogus"
# 403   (ובגוף: {"status":"error","code":"grant_invalid"})

curl -s "https://session-gateway.<account>.workers.dev/v1/poll?kind=approval&sessionCode=<קוד>&idNumber=<ת.ז.>"
# {"status":"error","message":"לא נמצא רישום"} לפני הרשמה, ואחריה {"status":"ok","approval":"waiting",…}
```

שתי בקשות ברצף לאותו סשן צריכות להחזיר את אותה תשובה בלי הרצת Apps Script שנייה — נבדק
בגיליון ההרצות: ‏`sessionSnapshot` אחד לכל ~2 שניות של סקר, לא אחד לכל נבחן.

## תקציב

100k בקשות/יום **לכל החשבון** — משותף עם ה‑TTS, ה‑proxy לתמונות ו‑worker הדו"חות.
בוקר בחינות של 40 נבחנים ≈ 18k בקשות; שלושה אתרים ≈ 55k (טעינת הבנק היא בקשה אחת למכשיר
לכל מבחן, לא לכל שאלה). מעל התקציב Cloudflare מחזיר 5xx, והלקוח נופל חזרה לסקר ישיר אחרי
3 כשלים רצופים.

## בדיקות

`node tests/gateway.test.cjs` (בלי רשת: fetch מזויף וסָפוּר, שעון מזויף, `env.ASSETS` בזיכרון).
`node tests/bank_invariants.test.cjs` (קורא את הבנק מתוך `assets/`).
`package.json` כאן קיים רק כדי ש‑`worker.js` ייטען כ‑ESM ב‑Node.
