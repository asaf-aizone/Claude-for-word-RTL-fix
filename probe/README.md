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

---

## Outlook host discovery (M0 for v0.3.0)

**המטרה:** לאמת את ההנחות שמופיעות ב-`docs/OUTLOOK-EXPANSION-PLAN.md` סעיף 2 לפני שכותבים שורת קוד production. אסור לעבור ל-M1 בלי תשובות מבוססות-ראיות ל-6 השאלות.

**מה ה-probe עונה עליו:**

1. Q1 - האם Claude ב-Outlook רץ בכלל ב-WebView2? (`msedgewebview2.exe` חייב להיות צאצא של `OUTLOOK.EXE` או `olk.exe`.)
2. Q2 - האם ה-URL זהה ל-Word/Excel/PowerPoint (`pivot.claude.ai`)?
3. Q3 - מה הערך המדויק של `_host_Info=` עבור Outlook?
4. Q4 - האם `--remote-debugging-port=0` נתפס על ידי ה-WebView2 של Outlook?
5. Q5 - מה ההבדל בין `OUTLOOK.EXE` הקלאסי לבין New Outlook (`olk.exe` / Appx)? תמיכה באחד, בשניהם, או באחד-משניהם תחילה?
6. Q6 - האם Outlook חולק `msedgewebview2.exe` host pool עם Word/Excel/PowerPoint?

**הרצה - Outlook הקלאסי:**

1. סגור את כל חלונות Outlook (הקלאסי **וגם** New Outlook). הסקריפט יסרב לרוץ אם אחד מהם פעיל - תהליך שכבר רץ לא יקלוט את ה-env var, וזה ייתן false negative ל-Q4.
2. דאבל-קליק על `outlook-host-discovery.bat`. הוא מפעיל את `OUTLOOK.EXE` הקלאסי עם `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0`. לא נכתב כלום ל-HKCU - ה-env var תקף רק לתהליך שהושק.
3. ב-Outlook: פתח מייל כלשהו, היכנס לתפריט Apps / Add-ins, פתח את Claude. חכה כמה שניות לטעינה.
4. בטרמינל נפרד:
   ```
   node outlook-host-discovery.js
   ```
5. ה-probe מדפיס שש סקציות, אחת לכל שאלה, עם הראיה (PID, port, URL, _host_Info). בסוף - שורת `GO` או `NO-GO`.

**איך לענות גם ל-Q6 (שיתוף host pool):**

הוסף Office app נוסף לפני הרצת ה-probe:
- פתח Word/Excel/PowerPoint דרך `launch-office-dynamic.bat` (אסור להפעיל מה-taskbar - לא יהיה env var).
- פתח את Claude add-in גם שם.
- אז הרץ `node outlook-host-discovery.js`. הוא יזהה את ה-PIDs של כל המארחים ויבדוק חפיפה.

**מה לעשות אם NO-GO:**

- Q1=NO - ייתכן ש-Outlook הקלאסי משתמש ב-IE legacy host או ב-stack שונה לתוסף. בלי WebView2 הארכיטקטורה הזו לא ישימה. תכנית ההרחבה מבוטלת או נדחית.
- Q2=NO ו-URL אחר - יש לבדוק האם המבנה ה-DOM זהה. אם זהה, התאמת `URL_PATTERN_PRIMARY` ב-injector מספיקה. אם DOM שונה, יש לכתוב סלקטור חדש.
- Q4=NO - WebView2 לא קלט את הדגל. בדוק שאכן הופעל דרך ה-bat ולא דרך taskbar. ב-New Outlook (Appx) זה צפוי - app container חוסם את ה-env var.

**ניקוי:** סגור Outlook (ואת שאר ה-Office אם פתחת לטובת Q6). אין שאריות במערכת.

### ממצאי הרצה ראשונה - Outlook קלאסי (2026-05-12)

הרצה על Windows 11 Pro 26200, Microsoft 365 גרסה 16.0.19929.20136. תוסף Claude נטען ידנית בתוך מייל פתוח.

| שאלה | תשובה | ראיה |
|------|-------|------|
| Q1 - WebView2? | **YES** | 15 process `msedgewebview2.exe` הם descendants של `OUTLOOK.EXE` (PID 21696). מהם שניים עם LISTENING TCP port |
| Q2 - אותו URL? | **YES** | `https://pivot.claude.ai/...` - אותה domain כמו Word/Excel/PowerPoint |
| Q3 - `_host_Info=` | **YES** | `Outlook$Win32$16.02$he-IL$$$$16`. ה-app name (קטע ראשון לפני `$`) הוא `Outlook` - מתאים לפורמט שלפיו `port-discovery.js` כבר מסווג. גרסת ה-Office API היא `16.02` (ב-Word היא `16.01`) |
| Q4 - `--remote-debugging-port=0`? | **YES** | port דינמי 54812 על PID 27704, עם Claude target זמין דרך `/json/list` |
| Q5 - Classic vs New | רק Classic נבדק בריצה זו. `olk.exe` עלה רגעית עם הפעלת ה-bat (כנראה כתהליך עזר של Classic) אך לא חשף debug port. New Outlook הופרד למסלול נפרד |
| Q6 - host pool משותף? | UNKNOWN - נדחה. הדגימה כללה רק Outlook |

תוספות שאינן בשאלות אבל קריטיות לעיצוב M1:

- **שני** WebView2 process descendants של OUTLOOK.EXE מאזינים על TCP בו-זמנית. אחד מארח את Claude (port 54812), השני (port 58973) מארח 3 עמודים שאינם Claude - ככל הנראה task pane של תוסף אחר או host shell. ה-port discovery הקיים מסנן לפי `URL_PATTERN_PRIMARY` ולכן יטפל בזה ללא שינוי.
- URL של Outlook מכיל פרמטר ייעודי `m=outlook-1.0.0.4` שאינו קיים ב-Word/Excel/PowerPoint. לא נדרש לזיהוי (המנגנון של `_host_Info=` מספיק), אך אפשר להשתמש בו כסיגנל משלים.
- locale נקבע ל-`he-IL` בכותרת ה-`_host_Info=`. ה-DOM הזמין להזרקה אמור להיות זהה ל-Word עברית.

**החלטת go/no-go:** **GO** ל-Outlook הקלאסי. שלוש ההנחות החוסמות (Q1, Q2, Q3) חיוביות; Q4 חיובי גם הוא, ולכן ארכיטקטורת ה-wrapper של v0.2.x עובדת as-is. M1 יכול להתחיל לפי תכנית סעיף 5 - הקובץ `lib/office-apps.js`, `outlook-wrapper.bat`, ופריט `Connect Outlook` ב-tray. הקשחות סעיף 4 (M2) חובה לפני release.

**מסלול New Outlook (`olk.exe`):** נדחה. הסיבה: Appx app container עלול לחסום `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`. נדרש probe נפרד שמפעיל את ה-Appx (לא ניתן להפעיל את ה-Appx מ-.bat עם env var בצורה הרגילה, צריך `Invoke-AppxPackage`/`explorer.exe shell:AppsFolder\...` או PowerShell דדיקטי). M0 שני, לא בסשן הזה.

**Q6 (host pool):** מומלץ לאמת לפני M1, אך לא חוסם. ההרצה דורשת לפתוח Outlook + Word במקביל (Word דרך `word-wrapper.bat` הקיים, Outlook דרך ה-probe.bat), ולהריץ את `outlook-host-discovery.js` שוב.

### ממצאי הרצה שנייה - Outlook + Word + Excel בו-זמנית (2026-05-12)

הרצה שנייה במטרה לענות על Q6, אבל חשפה ממצא בעדיפות גבוהה יותר.

הוספנו Word, Excel, PowerPoint להפעלה לצד Outlook שכבר היה פתוח. המטרה הייתה לבדוק האם Outlook חולק `msedgewebview2.exe` host pool עם Word/Excel/PowerPoint. בפועל - גילינו משהו אחר.

**Q6 - תשובה: לא חולק host pool.** הלוג של ה-injector הקיים (v0.2.2) מראה שלושה targets במקביל על שלושה פורטים ושלושה PIDs נפרדים:

```
Word@58973 ... | unknown@54812 [Outlook] | Excel@59652 ...
```

כל אפליקציית Office מקבלת process WebView2 משלה עם פורט debug נפרד. ה-injector מסתדר עם זה דרך URL filtering ב-`port-discovery.js`, כי הוא מסנן לפי `URL_PATTERN_PRIMARY` ולא לפי PID מארח.

#### ממצא בעדיפות גבוהה - silent CDP attach ל-Outlook ב-v0.2.2

**הראיה (`%TEMP%\claude-word-rtl.log`):**

```
2026-05-12T03:38:40.137Z  targets (tick 951): matched=1 entries=[unknown@54812
  https://pivot.claude.ai/?m=outlook-1.0.0.4&_host_Info=Outlook$Win32$16.02$he-IL$$$$16]
2026-05-12T03:38:40.936Z  Attached & injected: [unknown] port=54812 ... -> injected
```

ה-injector של v0.2.2 שרץ ברקע מ-Startup (PID 16344) **זיהה את ה-target של Outlook, ביצע attach דרך WebSocket ל-CDP, והזריק RTL CSS** - ללא לחיצת Connect Outlook (אין כזה ב-v0.2.2), ללא דיאלוג אזהרה, וללא שום אינדיקציה ב-tray (`apps.json` הקיים לא יודע על Outlook ולכן אין סטטוס label שמשקף את זה). הסיווג ב-log הוא `[unknown]` כי `lib/office-apps.js` לא מכיל ערך `Outlook`, אבל הסיווג רק שולט בתווית - לא בהחלטה האם להיצמד.

**מנגנון:** `scripts/port-discovery.js` סורק את כל ה-`msedgewebview2.exe` PIDs במערכת, מוצא LISTENING TCP, ובודק כל port מול `/json/list`. כל target ש-URL שלו מתאים ל-`pivot.claude.ai` (URL_PATTERN_PRIMARY) או ל-`*.claude.ai` (fallback) - מועבר ל-attach. אין בדיקה לפי שם אפליקציית Office. כל עוד ה-URL מתאים, ה-injector נצמד.

**מה גרם לזה לקרות בשיחה הזו:** ה-probe.bat שלנו הגדיר `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0` ב-scope של OUTLOOK.EXE. WebView2 של Outlook קלט את הדגל, פתח פורט debug דינמי. ה-injector של v0.2.2 שכבר רץ, סרק, מצא, התחבר.

**רמת חשיפה במציאות של v0.2.2:** בלי probe.bat או wrapper שמגדיר את ה-env var ל-OUTLOOK.EXE, Outlook לא חושף CDP. כלומר אצל משתמש v0.2.2 רגיל שלא הריץ קוד probe ייעודי, אין חשיפה כרגע. אבל הבעיה היא בדיוק עבור v0.3.0: ברגע שמוסיפים `outlook-wrapper.bat` (כמתוכנן ב-M1), כל הפעלת Outlook דרך פריט ה-tray "Connect Outlook" תפעיל את אותו flow בדיוק - ולא יהיו הקשחות סעיף 4 פעילות עד M2.

**השלכה לסדר העבודה ב-M1:**

התכנית המקורית מסעיף 7 הציעה:
- M1 - קוד מינימלי (`office-apps.js`, `outlook-wrapper.bat`, פריט tray), בלי הקשחות
- M2 - הקשחות אבטחה

הראיה כאן מצריכה היפוך:

- **M1 חייב לכלול לפחות הקשחה 4.3** (הקשחת ה-target filter ב-`scripts/inject.js` - blocklist על `_host_Info=Outlook$` כברירת מחדל, opt-in מפורש דרך apps.json/flag כדי להתיר). כל קוד אחר של M1 (wrapper, פריט tray, רישום ב-apps.json) חייב לבוא אחרי הוספת ה-filter, לא לפניו. אחרת חלון זמן ההפצה של v0.3.0 הופך את ה-attack surface שמתוארת בסעיף 3 לבעיה אקטיבית.
- **הקשחה 4.1** (manualOnly) חייבת להיכלל ב-M1 גם כן, מאותה סיבה - בלעדיה ה-tray ישגר את ה-injector אוטומטית כש-Outlook רץ (קיים היום, ב-`tray-icon.ps1` ב-30s cooldown), וה-injector ינסה attach.
- **הקשחות 4.2, 4.4, 4.5, 4.6** יכולות להישאר ב-M2.

**מסקנת go/no-go מעודכנת:** **GO** עם הסתייגות: סדר העבודה ב-M1 הופך - הקשחות 4.1 ו-4.3 לפני wrapper ו-tray menu. כל שינוי אחר ב-`OUTLOOK-EXPANSION-PLAN.md` סעיף 7 (אבני דרך) או סעיף 3 (איום) ייעשה בנפרד מסשן זה.

**מצב סוף הסשן:** קבצי probe חדשים בלבד נכתבו (`probe/outlook-host-discovery.bat`, `probe/outlook-host-discovery.js`) וסקציה זו נוספה ל-`probe/README.md`. שום שינוי ב-`lib/`, `scripts/`, או `docs/OUTLOOK-EXPANSION-PLAN.md`. לא בוצע commit לגיט.
