const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// Load logo as base64
const logoPath = path.join(__dirname, "logo.png");
const logoB64 = "image/png;base64," + fs.readFileSync(logoPath).toString("base64");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Vitaly Gitelman";
pres.title = "סמכות הרישוי צה״ל — מערכת מבחני תאוריה דיגיטלית";

// Beitar colors
const GOLD = "FFD130";
const GOLD_DIM = "3D3000";
const BLACK = "111111";
const DARK2 = "1A1A1A";
const DARK3 = "222222";
const WHITE = "FFFFFF";
const GRAY = "9CA3AF";
const GREEN = "4ADE80";
const RED = "F87171";

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.3 });

// ===================== SLIDE 1: TITLE =====================
{
  const slide = pres.addSlide();
  slide.background = { color: BLACK };

  // Subtle gold glow
  slide.addShape(pres.shapes.OVAL, { x: 2, y: -1, w: 6, h: 6, fill: { color: GOLD, transparency: 94 } });

  // Logo
  slide.addImage({ data: logoB64, x: 4.1, y: 0.4, w: 1.8, h: 1.8 });

  // Title
  slide.addText("סמכות הרישוי — צה״ל", {
    x: 0.5, y: 2.5, w: 9, h: 0.8,
    fontSize: 40, fontFace: "Arial", bold: true, color: WHITE, align: "center", rtlMode: true
  });
  slide.addText("מערכת מבחני תאוריה דיגיטלית", {
    x: 0.5, y: 3.2, w: 9, h: 0.7,
    fontSize: 32, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });

  // Badges row
  const badges = ["📱 מבוסס דפדפן", "🌐 7 שפות", "🔊 הקראה קולית", "🛡️ אבטחה מלאה", "📊 דוחות בזמן אמת"];
  const bw = 1.7;
  const startX = (10 - badges.length * bw - (badges.length - 1) * 0.1) / 2;
  badges.forEach((b, i) => {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: startX + i * (bw + 0.1), y: 4.3, w: bw, h: 0.45,
      fill: { color: GOLD, transparency: 85 }, line: { color: GOLD, width: 1, transparency: 60 }, rectRadius: 0.15
    });
    slide.addText(b, {
      x: startX + i * (bw + 0.1), y: 4.3, w: bw, h: 0.45,
      fontSize: 10, fontFace: "Arial", color: GOLD, align: "center", valign: "middle", rtlMode: true
    });
  });

  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 2: WHY DIGITAL? =====================
{
  const slide = pres.addSlide();
  slide.background = { color: BLACK };

  slide.addText("למה דיגיטלי?", {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 36, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });
  slide.addText("היתרונות של מערכת דיגיטלית מול מבחנים בחוברות נייר", {
    x: 0.5, y: 0.9, w: 9, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
  });

  // OLD column
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.4, w: 4.5, h: 4.0, fill: { color: "1A0505" }, line: { color: "3D1111", width: 1 } });
  slide.addText("📓 חוברות נייר )שיטה ישנה(", {
    x: 5.15, y: 1.45, w: 4.5, h: 0.4,
    fontSize: 14, fontFace: "Arial", bold: true, color: RED, align: "center", valign: "middle", rtlMode: true
  });

  const oldItems = [
    "שאלות לא מעודכנות",
    "כולם מקבלים אותן שאלות — קל להעתיק",
    "בדיקה ידנית — טעויות אנוש",
    "בזבוז נייר — מאות חוברות לפח",
    "אין משוב מפורט לנבחן",
    "אין נתונים סטטיסטיים",
    "אין תמיכה בשמע",
    "לוגיסטיקה מורכבת",
    "אין תרגול מובנה"
  ];
  const oldText = oldItems.map((t, i) => ({
    text: "✗  " + t,
    options: { breakLine: i < oldItems.length - 1, fontSize: 11, color: "D4A0A0", fontFace: "Arial" }
  }));
  slide.addText(oldText, { x: 5.35, y: 1.9, w: 4.1, h: 3.4, rtlMode: true, valign: "top", paraSpaceAfter: 4 });

  // NEW column
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.4, w: 4.5, h: 4.0, fill: { color: "1A1500" }, line: { color: GOLD, width: 1, transparency: 70 } });
  slide.addText("💻 מערכת דיגיטלית )שיטה חדשה(", {
    x: 0.35, y: 1.45, w: 4.5, h: 0.4,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle", rtlMode: true
  });

  const newItems = [
    "מאגר מעודכן תמיד",
    "מבחן ייחודי לכל נבחן — בלתי אפשרי להעתיק",
    "בדיקה אוטומטית ומדויקת",
    "חיסכון מלא בנייר — אפס הדפסות",
    "משוב מפורט בשפת הנבחן",
    "דוחות וסטטיסטיקות בזמן אמת",
    "הקראה קולית ב-7 שפות",
    "אפס לוגיסטיקה — עובד מהטלפון",
    "4 מצבי תרגול מובנים"
  ];
  const newText = newItems.map((t, i) => ({
    text: "✓  " + t,
    options: { breakLine: i < newItems.length - 1, fontSize: 11, color: GREEN, fontFace: "Arial" }
  }));
  slide.addText(newText, { x: 0.55, y: 1.9, w: 4.1, h: 3.4, rtlMode: true, valign: "top", paraSpaceAfter: 4 });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 3: TWO PRODUCTS =====================
{
  const slide = pres.addSlide();
  slide.background = { color: DARK2 };

  slide.addText("שני מוצרים, פלטפורמה אחת", {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 34, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });

  // Exam card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.2, w: 4.5, h: 4.1,
    fill: { color: DARK3 }, line: { color: GOLD, width: 1, transparency: 50 }, shadow: makeShadow()
  });
  slide.addText("📝", { x: 5.2, y: 1.3, w: 4.5, h: 0.5, fontSize: 32, align: "center" });
  slide.addText("מבחנים רשמיים", {
    x: 5.2, y: 1.8, w: 4.5, h: 0.4,
    fontSize: 22, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });
  slide.addText("דף בוחן + דף נבחן", {
    x: 5.2, y: 2.15, w: 4.5, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
  });
  const examFeats = [
    "סשן מבחן עם קוד QR",
    "פיקוח נבחנים בזמן אמת",
    "פסילה אוטומטית + ביטול פסילה",
    "דוחות PDF, WA, SMS",
    "דו\"ח מנהל אתר + מפקד",
    "עד 50 נבחנים בסשן"
  ];
  slide.addText(examFeats.map((t, i) => ({
    text: "✓  " + t,
    options: { breakLine: i < examFeats.length - 1, fontSize: 11, color: "D4D4D4", fontFace: "Arial" }
  })), { x: 5.5, y: 2.55, w: 4.0, h: 2.6, rtlMode: true, valign: "top", paraSpaceAfter: 5 });

  // Practice card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.2, w: 4.5, h: 4.1,
    fill: { color: DARK3 }, line: { color: GOLD, width: 1, transparency: 50 }, shadow: makeShadow()
  });
  slide.addText("📚", { x: 0.3, y: 1.3, w: 4.5, h: 0.5, fontSize: 32, align: "center" });
  slide.addText("תרגול ולמידה", {
    x: 0.3, y: 1.8, w: 4.5, h: 0.4,
    fontSize: 22, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });
  slide.addText("דף מורה + דף תלמיד", {
    x: 0.3, y: 2.15, w: 4.5, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
  });
  const practFeats = [
    "מבחן תרגול — סימולציה מלאה",
    "תרגול לפי נושא עם משוב מיידי",
    "כרטיסיות )Flashcards(",
    "חזרה מרווחת — למידה חכמה",
    "הקראה קולית מלאה",
    "ניהול כיתה + מעקב תלמידים"
  ];
  slide.addText(practFeats.map((t, i) => ({
    text: "✓  " + t,
    options: { breakLine: i < practFeats.length - 1, fontSize: 11, color: "D4D4D4", fontFace: "Arial" }
  })), { x: 0.6, y: 2.55, w: 4.0, h: 2.6, rtlMode: true, valign: "top", paraSpaceAfter: 5 });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 4: FEATURES (3x3 grid) =====================
{
  const slide = pres.addSlide();
  slide.background = { color: BLACK };

  slide.addText("יכולות מרכזיות", {
    x: 0.5, y: 0.2, w: 9, h: 0.6,
    fontSize: 34, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });

  const features = [
    { icon: "📱", title: "מבוסס דפדפן", desc: "ללא התקנה — טלפון, טאבלט, מחשב" },
    { icon: "🌐", title: "7 שפות", desc: "עברית, רוסית, ערבית, אנגלית, צרפתית, ספרדית, אמהרית" },
    { icon: "🔊", title: "הקראה קולית", desc: "בחירת קול )גבר/אישה( ומהירות" },
    { icon: "🛡️", title: "אבטחה מרובת שכבות", desc: "טוקן, הצפנה, זיהוי מעבר מסך" },
    { icon: "📊", title: "דוחות מתקדמים", desc: "PDF, ניתוח קטגוריות, מפת חום" },
    { icon: "⚡", title: "זמן אמת", desc: "עדכון מיידי — רישום, סיום, פסילה" },
    { icon: "🔄", title: "ביטול פסילה", desc: "הנבחן ממשיך מאותו מקום" },
    { icon: "🧠", title: "למידה חכמה", desc: "חזרה מרווחת — שאלות ממוקדות" },
    { icon: "📤", title: "שיתוף תוצאות", desc: "WhatsApp, SMS, קישור" }
  ];

  const cardW = 2.8, cardH = 1.3, gapX = 0.3, gapY = 0.2;
  const gridW = 3 * cardW + 2 * gapX;
  const offsetX = (10 - gridW) / 2;

  features.forEach((f, i) => {
    const col = 2 - (i % 3); // RTL: right to left
    const row = Math.floor(i / 3);
    const x = offsetX + col * (cardW + gapX);
    const y = 1.0 + row * (cardH + gapY);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: DARK3 }, line: { color: GOLD, width: 0.5, transparency: 70 }
    });
    slide.addText(f.icon, { x, y: y + 0.08, w: cardW, h: 0.35, fontSize: 22, align: "center" });
    slide.addText(f.title, {
      x, y: y + 0.42, w: cardW, h: 0.3,
      fontSize: 13, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
    });
    slide.addText(f.desc, {
      x: x + 0.1, y: y + 0.72, w: cardW - 0.2, h: 0.5,
      fontSize: 10, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 5: EXAM STRUCTURE TABLE =====================
{
  const slide = pres.addSlide();
  slide.background = { color: BLACK };

  slide.addText("מבנה המבחן", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 34, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });
  slide.addText("30 שאלות · 40 דקות · ציון עובר: 86% )26/30(", {
    x: 0.5, y: 0.85, w: 9, h: 0.35,
    fontSize: 14, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
  });

  const headerOpts = { fill: { color: GOLD_DIM }, color: GOLD, bold: true, fontSize: 12, fontFace: "Arial", align: "center", valign: "middle" };
  const cellOpts = { color: "D4D4D4", fontSize: 12, fontFace: "Arial", align: "center", valign: "middle" };
  const gradeOpts = { color: GOLD, bold: true, fontSize: 14, fontFace: "Arial", align: "center", valign: "middle" };

  const tableRows = [
    [
      { text: "סה\"כ", options: headerOpts },
      { text: "ספציפי", options: headerOpts },
      { text: "תמרורים", options: headerOpts },
      { text: "חוק", options: headerOpts },
      { text: "הכרת הרכב", options: headerOpts },
      { text: "בטיחות", options: headerOpts },
      { text: "תיאור", options: headerOpts },
      { text: "דרגה", options: headerOpts }
    ],
    [
      { text: "30", options: { ...cellOpts, bold: true } },
      { text: "—", options: cellOpts },
      { text: "9", options: cellOpts },
      { text: "7", options: cellOpts },
      { text: "7", options: cellOpts },
      { text: "7", options: cellOpts },
      { text: "רכב פרטי", options: cellOpts },
      { text: "B", options: gradeOpts }
    ],
    [
      { text: "30", options: { ...cellOpts, bold: true } },
      { text: "8", options: cellOpts },
      { text: "6", options: cellOpts },
      { text: "6", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "טרקטור", options: cellOpts },
      { text: "1", options: gradeOpts }
    ],
    [
      { text: "30", options: { ...cellOpts, bold: true } },
      { text: "10", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "משא קל", options: cellOpts },
      { text: "C1", options: gradeOpts }
    ],
    [
      { text: "30", options: { ...cellOpts, bold: true } },
      { text: "14", options: cellOpts },
      { text: "4", options: cellOpts },
      { text: "3", options: cellOpts },
      { text: "4", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "משא כבד / גורר", options: cellOpts },
      { text: "C+E", options: gradeOpts }
    ],
    [
      { text: "30", options: { ...cellOpts, bold: true } },
      { text: "15", options: cellOpts },
      { text: "4", options: cellOpts },
      { text: "5", options: cellOpts },
      { text: "2", options: cellOpts },
      { text: "4", options: cellOpts },
      { text: "אוטובוס / מונית", options: cellOpts },
      { text: "D", options: gradeOpts }
    ]
  ];

  slide.addTable(tableRows, {
    x: 0.5, y: 1.4, w: 9,
    colW: [0.8, 0.8, 0.9, 0.8, 1.1, 0.9, 1.6, 1.0],
    rowH: [0.5, 0.45, 0.45, 0.45, 0.45, 0.45],
    border: { pt: 0.5, color: "333333" },
    fill: { color: DARK3 }
  });

  slide.addText("שאלות נבחרות אקראית מכל קטגוריה · סדר תשובות מעורבב · שאלות \"זכות קדימה\" נכללות בכל הדרגות", {
    x: 0.5, y: 4.5, w: 9, h: 0.4,
    fontSize: 11, fontFace: "Arial", color: "666666", align: "center", rtlMode: true
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 6: STATS =====================
{
  const slide = pres.addSlide();
  slide.background = { color: DARK2 };

  slide.addText("המערכת במספרים", {
    x: 0.5, y: 0.5, w: 9, h: 0.7,
    fontSize: 36, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });

  const stats = [
    { num: "1,500+", label: "שאלות במאגר" },
    { num: "5", label: "דרגות רישיון" },
    { num: "7", label: "שפות נתמכות" },
    { num: "4", label: "מצבי תרגול" }
  ];

  const statW = 2.1, statGap = 0.2;
  const statTotalW = stats.length * statW + (stats.length - 1) * statGap;
  const statStartX = (10 - statTotalW) / 2;

  stats.forEach((s, i) => {
    const x = statStartX + i * (statW + statGap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.8, w: statW, h: 2.2,
      fill: { color: DARK3 }, line: { color: GOLD, width: 1, transparency: 60 }
    });
    slide.addText(s.num, {
      x, y: 2.0, w: statW, h: 1.2,
      fontSize: 40, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle"
    });
    slide.addText(s.label, {
      x, y: 3.1, w: statW, h: 0.6,
      fontSize: 14, fontFace: "Arial", color: GRAY, align: "center", valign: "middle", rtlMode: true
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 7: LANGUAGES =====================
{
  const slide = pres.addSlide();
  slide.background = { color: BLACK };

  slide.addText("תמיכה ב-7 שפות", {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 34, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });
  slide.addText("כל שאלה מתורגמת באופן מקצועי — כולל מונחים טכניים", {
    x: 0.5, y: 0.9, w: 9, h: 0.35,
    fontSize: 14, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
  });

  const langs = [
    { code: "IL", name: "עברית" },
    { code: "RU", name: "רוסית" },
    { code: "SA", name: "ערבית" },
    { code: "GB", name: "אנגלית" },
    { code: "FR", name: "צרפתית" },
    { code: "ES", name: "ספרדית" },
    { code: "ET", name: "אמהרית" }
  ];

  const langW = 1.15, langH = 1.6, langGap = 0.13;
  const langTotalW = langs.length * langW + (langs.length - 1) * langGap;
  const langStartX = (10 - langTotalW) / 2;
  const langY = 1.8;

  langs.forEach((l, i) => {
    const x = langStartX + i * (langW + langGap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: langY, w: langW, h: langH,
      fill: { color: DARK3 }, line: { color: GOLD, width: 0.5, transparency: 70 }
    });
    slide.addText(l.code, {
      x, y: langY + 0.15, w: langW, h: 0.8,
      fontSize: 28, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle"
    });
    slide.addText(l.name, {
      x, y: langY + 1.0, w: langW, h: 0.4,
      fontSize: 13, fontFace: "Arial", bold: true, color: WHITE, align: "center", rtlMode: true
    });
  });

  // Additional info below cards
  slide.addText("כל ממשק — שאלות, תשובות, הקראה קולית ומשוב — זמין בכל שפה", {
    x: 0.5, y: 3.7, w: 9, h: 0.4,
    fontSize: 12, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 8: SECURITY =====================
{
  const slide = pres.addSlide();
  slide.background = { color: DARK2 };

  slide.addText("אבטחה ומניעת רמאות", {
    x: 0.5, y: 0.2, w: 9, h: 0.6,
    fontSize: 34, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });

  const secCards = [
    {
      title: "🛡️ צד שרת",
      items: ["טוקן אימות — תוקף 12 שעות", "סיסמה מוצפנת + נעילה", "חישוב ציון בשרת בלבד", "הגבלת 50 נבחנים"]
    },
    {
      title: "📱 אנטי-צ'יט בטלפון",
      items: ["מעבר אפליקציה — 3 אזהרות", "שיחה 5+ שניות = פסילה", "פיצול מסך )65% >(", "התעלמות מניתוקים < 2 שניות"]
    },
    {
      title: "🖥️ אנטי-צ'יט מחמיר במחשב",
      items: ["מסך מלא חובה — יציאה = פסילה", "Alt+Tab = פסילה תוך 2ש' )ללא אזהרות(", "Snap / שינוי גודל )90% >(", "0 ms סובלנות ל-blur"]
    },
    {
      title: "🔄 התאוששות חכמה + בקרה",
      items: ["ביטול פסילה — המשך מאותו מקום", "שמירת מצב מבחן + איפוס", "דשבורד מפקד + מנהל אתר", "מפת חום + ניתוח נושאים"]
    }
  ];

  secCards.forEach((card, i) => {
    const col = i < 2 ? 1 : 0;
    const row = i % 2;
    const x = 0.3 + (1 - col) * 4.85;
    const y = 1.0 + row * 2.15;
    const cw = 4.55, ch = 1.95;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cw, h: ch,
      fill: { color: DARK3 }, line: { color: GOLD, width: 0.5, transparency: 70 }
    });

    slide.addText(card.title, {
      x: x + 0.15, y: y + 0.08, w: cw - 0.3, h: 0.35,
      fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "right", rtlMode: true, margin: 0
    });

    slide.addText(card.items.map((t, j) => ({
      text: "✓  " + t,
      options: { breakLine: j < card.items.length - 1, fontSize: 11, color: "B0B0B0", fontFace: "Arial" }
    })), { x: x + 0.2, y: y + 0.45, w: cw - 0.4, h: 1.4, rtlMode: true, valign: "top", paraSpaceAfter: 4 });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 9: CLASSROOM DEPLOYMENT =====================
{
  const slide = pres.addSlide();
  slide.background = { color: BLACK };

  slide.addText("פריסה בכיתת מחשבים", {
    x: 0.5, y: 0.2, w: 9, h: 0.6,
    fontSize: 34, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });
  slide.addText("התקנה אוטומטית של קיצורי דרך ל-30+ מחשבי כיתה — במצב אפליקציה ייעודי", {
    x: 0.5, y: 0.85, w: 9, h: 0.35,
    fontSize: 13, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
  });

  const depCards = [
    {
      title: "📦 חבילת התקנה",
      items: ["install-teacher.bat — 4 קיצורים", "install-student.bat — 2 קיצורים", "הורדת אייקונים אוטומטית + ICO", "Edge / Chrome --app=URL"]
    },
    {
      title: "✨ מצב אפליקציה",
      items: ["חלון נקי — ללא שורת כתובת", "ללא טאבים — מונע גלישה מקבילה", "אכיפת מסך מלא חזקה", "שמות בעברית: בוחן/מורה/נבחן/תרגול"]
    },
    {
      title: "🖥️ אכיפה מחמירה",
      items: ["זיהוי אוטומטי של סוג המכשיר", "מסך מלא חובה מתחילת המבחן", "Alt+Tab = פסילה תוך 2 שניות", "הוראות מותאמות PC לנבחן"]
    },
    {
      title: "🚀 פריסה ל-30 מחשבים",
      items: ["העתקת תיקיית deployment/", "הרצת BAT פעם אחת", "אפשרות פריסה דרך GPO/PsExec", "אין צורך בהתקנת תוכנות נוספות"]
    }
  ];

  depCards.forEach((card, i) => {
    const col = i < 2 ? 1 : 0;
    const row = i % 2;
    const x = 0.3 + (1 - col) * 4.85;
    const y = 1.35 + row * 1.95;
    const cw = 4.55, ch = 1.8;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cw, h: ch,
      fill: { color: DARK3 }, line: { color: GOLD, width: 0.5, transparency: 70 }
    });

    slide.addText(card.title, {
      x: x + 0.15, y: y + 0.08, w: cw - 0.3, h: 0.35,
      fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "right", rtlMode: true, margin: 0
    });

    slide.addText(card.items.map((t, j) => ({
      text: "✓  " + t,
      options: { breakLine: j < card.items.length - 1, fontSize: 11, color: "B0B0B0", fontFace: "Arial" }
    })), { x: x + 0.2, y: y + 0.45, w: cw - 0.4, h: 1.3, rtlMode: true, valign: "top", paraSpaceAfter: 4 });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// ===================== SLIDE 10: SUMMARY =====================
{
  const slide = pres.addSlide();
  slide.background = { color: BLACK };

  slide.addShape(pres.shapes.OVAL, { x: 2.5, y: -0.5, w: 5, h: 5, fill: { color: GOLD, transparency: 95 } });

  slide.addImage({ data: logoB64, x: 4.25, y: 0.3, w: 1.5, h: 1.5 });

  slide.addText("סמכות הרישוי — צה״ל", {
    x: 0.5, y: 2.0, w: 9, h: 0.6,
    fontSize: 30, fontFace: "Arial", bold: true, color: WHITE, align: "center", rtlMode: true
  });
  slide.addText("מערכת מבחני תאוריה דיגיטלית", {
    x: 0.5, y: 2.5, w: 9, h: 0.6,
    fontSize: 26, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
  });

  // Summary cards
  const sumCards = [
    { icon: "📝", title: "מבחנים רשמיים", desc: "בוחן + נבחן\nאבטחה מלאה" },
    { icon: "📚", title: "תרגול ולמידה", desc: "מורה + תלמיד\n4 מצבי למידה" },
    { icon: "📊", title: "ניהול ובקרה", desc: "דשבורד מפקד\nסטטיסטיקות" }
  ];

  const scW = 2.6, scGap = 0.3;
  const scTotalW = 3 * scW + 2 * scGap;
  const scStartX = (10 - scTotalW) / 2;

  sumCards.forEach((c, i) => {
    const x = scStartX + i * (scW + scGap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 3.3, w: scW, h: 1.5,
      fill: { color: DARK3 }, line: { color: GOLD, width: 1, transparency: 50 }
    });
    slide.addText(c.icon, { x, y: 3.35, w: scW, h: 0.4, fontSize: 24, align: "center" });
    slide.addText(c.title, {
      x, y: 3.72, w: scW, h: 0.3,
      fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "center", rtlMode: true
    });
    slide.addText(c.desc, {
      x, y: 4.05, w: scW, h: 0.65,
      fontSize: 11, fontFace: "Arial", color: GRAY, align: "center", rtlMode: true
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: GOLD } });
}

// Save
const outPath = path.join(__dirname, "product_presentation.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("Created: " + outPath);
}).catch(err => {
  console.error("Error:", err);
});
