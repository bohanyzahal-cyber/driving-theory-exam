# פריסת ה-Image Proxy Worker — צעד אחר צעד

זמן: ~10 דקות. עלות: 0₪.

## שלב 1: כניסה ל-Cloudflare Dashboard
1. פתח [https://dash.cloudflare.com](https://dash.cloudflare.com) — התחבר עם אותו חשבון של ה-TTS
2. בתפריט הצד השמאלי: **Workers & Pages**
3. לחץ **Create application** → **Create Worker**

## שלב 2: יצירת ה-Worker
1. **שם ה-Worker:** `images-proxy` (אם תפוס — נסה `image-proxy` או `theory-images`)
2. הכתובת תיבנה אוטומטית: `https://images-proxy.<accountname>.workers.dev`
3. לחץ **Deploy** (ה-Worker יקום עם קוד דמה)

## שלב 3: הדבקת הקוד
1. אחרי הפריסה, לחץ **Edit code**
2. **מחק את כל הקוד** שיש שם
3. פתח את הקובץ `deployment/image-proxy-worker.js` במחשב
4. **העתק את כל התוכן והדבק** בעורך של Cloudflare
5. לחץ **Save and deploy** (כפתור כתום למעלה מימין)

## שלב 4: בדיקה ראשונית
פתח בדפדפן:
```
https://images-proxy.<accountname>.workers.dev/health
```
אמור להופיע: `Image Proxy OK`

עכשיו נסה תמונה אמיתית:
```
https://images-proxy.<accountname>.workers.dev/img/BlobFolder/generalpage/tq_pic_01/he/TQ_PIC_31276.jpg
```
אמורה לעלות תמונה של תמרור או סיטואציית נהיגה.

ב-DevTools (F12 → Network) תראה את ה-header `X-Proxy-Cache: MISS` בפעם הראשונה ו-`HIT` בפעם השנייה.

## שלב 5: עדכון `examinee.html`
ב-`examinee.html` שורה 463 יש:
```js
var IMAGE_PROXY_BASE = 'https://images-proxy.bohanyzahal.workers.dev';
```

**אם בחרת שם אחר ל-Worker או חשבון אחר** — שנה את ה-URL בהתאם.

לאחר השינוי, commit + push:
```bash
git add examinee.html
git commit -m "עדכון URL של ה-image proxy"
git push
```

## שלב 6: בדיקה מקצה לקצה
1. חכה 1-2 דקות לסנכרון GitHub Pages
2. פתח את דף הנבחן בטלפון
3. הירשם למבחן דמה
4. בדוק שהתמונות עולות מהר ובלי תקלות

ב-DevTools במחשב — Network tab — תראה שכל הבקשות הולכות לדומיין של Worker, לא ל-gov.il.

## איך אפשר לדעת שזה עובד באמת?

### בקצר טווח
- אין "תמונה לא נטענה" בדפי הנבחן
- מחיקת קאש ופתיחה מחדש: התמונה עולה מיד (פעם שנייה ואילך מהאדג' של Cloudflare)

### בארוך טווח
- בדשבורד של ה-Worker: **Analytics → Requests** מראה כמה בקשות נכנסו, אחוז הצלחה, ו-latency ממוצע
- אם תראה `Requests = 6,000/day` ו-`Errors < 1%` — מעולה

## מה לעשות אם משהו לא עובד?

| תופעה | פתרון |
|---|---|
| `health` עובד אבל תמונות מחזירות 502 | ייתכן ש-gov.il חוסם זמנית. נסה שוב בעוד דקה |
| `400 Invalid path` | בדוק שה-URL שאתה שולח מתחיל ב-`/img/BlobFolder/...` |
| בדף הנבחן עדיין רואים שגיאות | הקאש של הדפדפן לא הסתנכרן. Ctrl+Shift+R / מחיקת קאש |
| תמונות עולות אבל איטיות | פעם ראשונה לכל תמונה זה לוקח ~500ms (הוצאה מ-gov.il). אחרי זה זה <50ms |

## שינוי URL בעתיד

אם תרצה לשנות את ה-Worker לשם אחר או חשבון אחר:
1. עדכן `IMAGE_PROXY_BASE` ב-`examinee.html`
2. **לבטל את הקאש בדפדפן** של הנבחנים (Ctrl+Shift+R)
3. להשאיר את ה-Worker הישן רץ עוד שבוע למקרה של תקלה

## צריכת המכסה החינמית

Cloudflare Workers — מכסה: **100,000 בקשות/יום חינם**.

| תרחיש | בקשות צפויות | % מהמכסה |
|---|---|---|
| 100 נבחנים × 20 תמונות | 2,000/יום | 2% |
| 500 נבחנים × 20 תמונות | 10,000/יום | 10% |
| 3,000 נבחנים × 20 תמונות | 60,000/יום | 60% |
| 5,000 נבחנים × 20 תמונות | 100,000/יום | 100% (אז $5/חודש קבוע) |

זה השימוש *הראשון* של תמונה. אחרי הפעם הראשונה, הקאש מטפל — לא נספרים בקשות חדשות.

**שורה תחתונה:** עד 5,000 נבחנים ביום — חינמי. למעלה — $5/חודש קבוע.
