# Probe - Office add-in diagnostics

סקריפטים לבדיקה חד-פעמית: האם Claude add-in נטען ב-Excel ו-PowerPoint, באיזה URL, והאם ה-pattern הקיים של ה-injector יתפוס אותו.

## הרצה

**חשוב:** לפני כל בדיקה, סגור את כל החלונות של האפליקציה הרלוונטית (Excel או PowerPoint). התוסף הנוכחי של Word יכול להישאר פתוח.

### Excel

1. סגור את כל חלונות Excel.
2. הפעל דאבל-קליק על `probe-excel.bat` (או מטרמינל: `probe-excel.bat`).
3. Excel ייפתח. הפעל ידנית את Claude add-in.
4. פתח Command Prompt בתיקייה הזו והרץ:
   ```
   node probe.js 9223
   ```
5. שלח את הפלט חזרה.

### PowerPoint

1. סגור את כל חלונות PowerPoint.
2. הפעל `probe-powerpoint.bat`.
3. PowerPoint ייפתח. הפעל ידנית את Claude add-in.
4. הרץ:
   ```
   node probe.js 9224
   ```
5. שלח את הפלט חזרה.

### בדיקת שלושתם ביחד (אופציונלי)

אם כבר יש לך Word פתוח עם התוסף הרגיל (port 9222), אפשר להריץ:
```
node probe.js
```
וזה יבדוק 9222, 9223 ו-9224 ביחד.

## מה אנחנו מחפשים בפלט

- **`[MATCH: pivot.claude.ai]`** - מעולה, אותו URL כמו ב-Word, ה-injector יעבוד כמו שהוא.
- **`[MATCH: claude.ai (fallback)]`** - טוב, ייתפס ע"י ה-fallback pattern.
- **`[no match]`** אבל ה-URL נראה קשור ל-Claude - יידרש עדכון של ה-pattern.
- **`[not reachable]`** - WebView2 debug port לא פתוח. ייתכן ש:
  - האפליקציה לא רצה עם משתנה הסביבה (בדוק שהפעלת דרך ה-.bat).
  - אין Claude add-in מותקן לאפליקציה הזו.
- **`[no targets]`** - ה-port מאזין אבל אין חלון WebView2 פתוח - הפעל את ה-add-in ונסה שוב.

## ניקוי

אחרי הבדיקה, פשוט סגור את Excel/PowerPoint. אין שאריות - ה-.bat לא משנה כלום במערכת, רק מעביר משתנה סביבה חד-פעמי לתהליך.

---

## POC - Dynamic ports (לאימות הארכיטקטורה החדשה)

**המטרה:** לאמת את חלופה B מסעיף 3.5 ב-OFFICE-EXPANSION-PLAN.md - האם `--remote-debugging-port=0` באמת גורם לכל אפליקציית Office לקבל פורט פנוי משלה, ואם ה-injector יכול לגלות את הפורטים אוטומטית דרך `tasklist`+`netstat`.

**שלבים:**

1. **סגור את כל אפליקציות Office** (Word, Excel, PowerPoint).

2. **הפעל את כולן עם הדגל הדינמי:**
   ```
   launch-office-dynamic.bat
   ```
   זה פותח את שלושתן עם `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0`.
   ה-env var תקף **רק לתהליכים שהושקו מהסקריפט** - לא נכתב ל-HKCU.

3. **פתח את Claude add-in בכל אחת מהשלוש** (ידנית, מתפריט ה-Add-ins).

4. **הרץ את סקריפט הגילוי:**
   ```
   node dynamic-port-discovery.js
   ```

5. **מה לבדוק בפלט:**
   - `Found N WebView2 PID(s)` - צריך להיות לפחות 3 (אחד לכל אפליקציה).
   - `Candidate (PID, port) pairs` - אמור להראות מיפוי PID->port לכל תהליך WebView2.
   - בסוף: `Apps detected: Word, Excel, Powerpoint` ו-`SUCCESS`.

6. **אם יש כשל:**
   - אם 0 WebView2 PIDs - האפליקציות לא הושקו עם ה-env var. ודא שהשתמשת ב-`launch-office-dynamic.bat` ולא פתחת Office דרך taskbar/double-click רגיל.
   - אם WebView2 PIDs מופיעים אבל בלי portים - הדגל לא נקלט. ייתכן בעיית גרסת WebView2 ישנה (פחות שכיח).
   - אם Apps זוהו פחות מ-3 - בדוק שה-Claude add-in פתוח בכל השלוש.

**ניקוי:** סגור את אפליקציות Office. אין שאריות במערכת.

**אם ה-POC מצליח:** ממשיכים לשלב M1 בתכנית - מימוש port discovery ב-`scripts/inject.js` כתחליף ל-`const PORT = 9222`.

**אם ה-POC נכשל:** חוזרים לחלופה D המקורית (3 wrappers + file associations). מתעדים את הסיבה ב-`docs/bugs/` חדש.
