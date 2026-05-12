# Claude for Office RTL Fix - תוכנית הרחבה ל-Outlook

**סטטוס:** טיוטה לביצוע בסשן חדש
**גרסת יעד:** v0.3.0
**מסמך מקור:** [OFFICE-EXPANSION-PLAN.md](OFFICE-EXPANSION-PLAN.md) (v0.2.0, הרחבה דומה אך לא זהה)
**תאריך:** 2026-05-12

---

## 1. החזון

הוספת תמיכת RTL ל-Claude add-in ב-Outlook desktop לפי אותו דפוס של v0.2.0:
פריט "Connect Outlook" בתפריט ה-tray, סטטוס per-app חדש "Outlook: connected/...",
ו-wrapper נפרד `outlook-wrapper.bat` שמפעיל את WebView2 עם debug-port דינמי.

ההבדל הקריטי מ-Word/Excel/PowerPoint: **תוכן המייל הפעיל הוא PII רגיש**.
הפתרון יקבל **מצב אבטחה מוקשח** ייעודי ל-Outlook (סעיף 4) שלא הופעל על שאר האפליקציות.

---

## 2. הנחות שצריך לאמת ב-probe (M0) לפני שכותבים קוד

המסלול של v0.2.0 התחיל ב-POC שאישר ש-Word/Excel/PowerPoint כולם משתמשים באותה shell של pivot.claude.ai דרך WebView2. עבור Outlook אסור להניח את זה. בלי probe נכשל.

### שאלות שה-probe חייב לענות עליהן

1. **האם Claude ב-Outlook רץ ב-WebView2 בכלל?** Outlook (במיוחד "New Outlook") עבר ב-2024-2025 ל-stack שונה מבוסס Edge/WebView2 אבל לא בכל הגרסאות. צריך לוודא ש-`msedgewebview2.exe` הוא תהליך הילד של `OUTLOOK.EXE`. אם זה לא ה-case (למשל אם המארח הוא IE legacy בגרסה ישנה, או PWA browser) - **התכנית הזו לא ישימה.**
2. **האם ה-URL זהה?** האם זה גם `https://pivot.claude.ai/` או URL אחר? אם זהה - INJECTOR_SCRIPT ו-RTL_CSS עובדים as-is. אם שונה - יש לבדוק שהמבנה ה-DOM זהה.
3. **מה ה-`_host_Info=`?** משוער `_host_Info=Outlook$Win32$16.01$...` אבל ייתכן שונה. צריך לתעד את הערך המדויק.
4. **האם debug-port דינמי (`--remote-debugging-port=0`) נתפס ע"י Outlook?** WebView2 צריך לקרוא את `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`. כל גרסת Outlook קלסי תעשה את זה, אבל New Outlook (Monarch) ייתכן ומוגן יותר.
5. **האם New Outlook ו-Outlook הקלאסי שונים?** ייתכן שיש שני executables (`OUTLOOK.EXE` ל-Win32 ו-`olk.exe` או UWP ל-New Outlook). ה-probe יבדוק את שניהם.
6. **קבוצות מארח משותפות:** ב-v0.2.0 התגלה ש-Word ו-PowerPoint חולקים אותו תהליך WebView2. צריך לבדוק אם Outlook מצטרף לאותו host pool (פותחים את כל הארבעה ובודקים PIDs ב-`tasklist`+`netstat`).

### תוצרי probe

- `probe/outlook-host-discovery.bat` ו-`probe/outlook-host-discovery.js` (אנלוגי ל-`probe/dynamic-port-discovery.js`).
- דוח קצר ב-`probe/README.md` שמסכם תשובות 1-6.
- **gate decision:** רק אם 1-3 חיוביים ממשיכים ל-M1.

---

## 3. שיקולי אבטחה - הסיבה שזה לא אותו שינוי מינורי

### האיום הקיים בכל v0.2.x ולמה הוא קריטי יותר עבור Outlook

`scripts/inject.js` פותח WebSocket ל-CDP על port דינמי localhost. CDP הוא surface מלא: כל תהליך מקומי שיודע למצוא את הפורט יכול להתחבר ולקרוא את כל ה-DOM של ה-panel, להזריק JS משלו, ולקרוא state של JavaScript בתוך ה-page.

עבור Word/Excel/PowerPoint, מה ש-Claude רואה ב-panel זה מה שהמשתמש העתיק/הצביע/הצמיד בכוונה. עבור **Outlook**, ברגע שהמשתמש לוחץ "Summarize this email" או "Draft a reply" - **תוכן המייל המלא נכנס ל-DOM של ה-panel**, ולכן חשוף ל-CDP.

### וקטור התקפה ריאלי

תוכנה מקומית (badware/PUA, browser extension עם access ל-localhost, script שרץ תחת אותו משתמש) ש:
1. סורקת `tasklist` ומוצאת `msedgewebview2.exe` (פתוח ב-Office, Teams, WidgetService וכו').
2. `netstat -ano` ומוצאת LISTENING ports של אותם PIDs.
3. שולחת GET ל-`http://localhost:<port>/json/list` ומסננת לפי `pivot.claude.ai`.
4. מתחברת ב-WebSocket ומריצה `Runtime.evaluate({ expression: 'document.body.innerText' })`.

ב-Word זה ייתן טיוטות שהמשתמש העתיק ל-Claude. **ב-Outlook זה ייתן כל מייל שהמשתמש סיכם/ענה עליו בסשן הנוכחי.**

### למה זה לא מתויק כ"החלטה לא לטפל"

ה-attack surface קיים גם היום עבור Word. הסיבות שלא נפתר עד היום:
- localhost-only, צריך תהליך מקומי (כלומר machine compromise כבר קיים)
- חלון זמן: פורט פתוח רק כשה-Office app רץ עם debug-port
- אין persistent state - לא נשמר tokens/credentials, רק תוכן UI הרגעי

עבור Outlook, הסיכון משתנה איכותית: מיילים יכולים להכיל סודות מסחריים, אשראי, מסמכים משפטיים, סיסמאות זמניות, MFA codes. גם **חלון של 30 שניות עם CDP פתוח על תוכן מייל הוא יותר רגיש מ-30 שניות על מסמך Word**.

### עדכון מ-M0 (2026-05-12) - האיום אומת בפועל

בזמן הרצת ה-probe הקיים `probe/outlook-host-discovery.js` נצפה שה-injector של v0.2.2 (שרץ ברקע מ-Startup) **ביצע attach + inject ל-target של Outlook אוטומטית, ללא לחיצת Connect Outlook וללא דיאלוג**. הראיה ב-`%TEMP%\claude-word-rtl.log`:

```
2026-05-12T03:38:40  matched=1 entries=[unknown@54812
  https://pivot.claude.ai/?m=outlook-1.0.0.4&_host_Info=Outlook$Win32$16.02$he-IL$$$$16]
2026-05-12T03:38:40  Attached & injected: [unknown] port=54812 ... -> injected
```

`scripts/port-discovery.js` סורק כל `msedgewebview2.exe`, ו-`scripts/inject.js` נצמד לכל target עם URL שמתאים ל-`pivot.claude.ai` ללא בדיקת זהות אפליקציית Office. המשמעות: ברגע שמישהו יפתח את Outlook עם debug port דינמי (היום: רק דרך ה-probe.bat; מחר: דרך `outlook-wrapper.bat` המתוכנן ל-M1), ה-injector הקיים יחבר אוטומטית ויעמוד פתוח כל זמן ש-Outlook רץ. אצל משתמש v0.2.2 רגיל זה עוד לא קורה כי אין wrapper, אבל ברגע ש-v0.3.0 יישוחרר - זה יקרה לכל מי שמשדרג, בלי שום ניראות.

זה לא איום עתידי. זה fix של בעיה קיימת. ההקשחות בסעיף 4 הן תנאי הכרחי לפני שחרור, לא תוספת איכות.

לתיעוד מלא של ההרצה והממצאים: `probe/README.md` תחת "ממצאי הרצה שנייה".

---

## 4. ההקשחות שיתווספו ל-Outlook (ולא חלות אוטומטית על השאר)

### 4.1 Opt-in מפורש כל הפעלה - אין auto-launch

ב-Word/Excel/PowerPoint, אם ה-injector קרס בזמן ש-Office רץ, ה-tray מפעיל מחדש את ה-injector אוטומטית (cooldown 30s). **עבור Outlook זה כבוי**. הסיבה: כל הפעלת CDP חייבת להיות מודעת ומפורשת. אם המשתמש לא לחץ Connect Outlook באקטיביות, אסור לפתוח לו פורט debug.

מימוש: דגל ב-`apps.json` שמסמן `manualOnly: true` עבור Outlook, או צ'ק ב-`tray-icon.ps1` שלא מאוטו-לאנץ' כש-`OUTLOOK.EXE` רץ.

### 4.2 דיאלוג אזהרה לפני כל Connect Outlook

`tray-icon.ps1` יציג דיאלוג OK/Cancel בעברית ובאנגלית לפני שמפעיל את `outlook-wrapper.bat`:

> **חיבור CDP ל-Outlook חושף את תוכן המייל הפעיל לכל תהליך מקומי שרץ תחת המשתמש שלך.** הסיכון הוא רק מתוכנות שכבר רצות במחשב; localhost לא נגיש מבחוץ. אם יש EDR ארגוני, הוא עלול לדווח על הפעולה. רוצה להמשיך?

הדיאלוג מציג checkbox "אל תשאל אותי שוב" שנשמר לסשן הנוכחי בלבד (בזיכרון של ה-tray, לא ב-registry).

### 4.3 הקשחת ה-target filter

`scripts/port-discovery.js` ו-`scripts/inject.js` יהדקו את ה-pattern:
- **רק** `pivot.claude.ai` (לא `*.claude.ai` כ-fallback). זה כבר primary היום אבל ה-fallback קיים. עבור Outlook נסיר לחלוטין.
- בנוסף, validate שה-target URL מכיל `_host_Info=Outlook$` בדיוק. אם לא, **detach מיידי**. זה מגן מפני תרחיש שבו תהליך זדוני יוצר WebView2 משלו שמתחזה ל-pivot.claude.ai עם host_Info מזויף - לא תוקף ריאלי גבוה, אבל זול.

### 4.4 Auto-disconnect אחרי N דקות

`scripts/inject.js` יחזיק ספירה לאחור per-target עבור Outlook (לא עבור Word/Excel/PowerPoint). אחרי 15 דקות של פעילות הזרקה רציפה, ה-injector מתנתק מה-target ומדווח `DISCONNECTED` ל-`apps.json`. המשתמש צריך ללחוץ Connect Outlook שוב כדי לחדש.

ההנמקה: גם אם המשתמש פתח Outlook לבוקר עבודה, אין סיבה שה-CDP יהיה פתוח רציף 8 שעות. סיכון מצטבר.

### 4.5 לוגינג מצומצם

`scripts/inject.js` מתעד היום את ה-URL המלא של כל target ב-`%TEMP%\claude-word-rtl.log`. ה-URL מכיל את ה-`et=...` parameter שהוא base64 של מטא-דאטה (account ID, tenant ID, expiry). **עבור Outlook נעטוף את ה-URL ב-redaction** - נשאיר רק `https://pivot.claude.ai/?...&_host_Info=Outlook$...` ונחתוך את ה-`et=` וכל פרמטר שאינו בייט-ליטרל ידוע. הלוג עדיין שמיש לדיבוג חיבור אבל לא חושף tenant ID לכל מי שמגיע ל-`%TEMP%`.

### 4.6 הסרה מהירה - "Disconnect Outlook only"

תפריט ה-tray היום מציע "Disconnect all" שסוגר את כל ה-Office. הוספה: **"Disconnect Outlook only"** שמשאיר את Word/Excel/PowerPoint מחוברים אבל מנתק את Outlook ספציפית. זה דורש שינוי ב-`inject.js` כדי לתמוך detach per-app, מה שלא קיים היום.

---

## 5. שינויי קוד מתוכננים

### תוספות

| קובץ | שינוי | סדר גודל |
|------|--------|----------|
| `lib/office-apps.js` | להוסיף `{ name: 'Outlook', processName: 'OUTLOOK.EXE', urlHostInfo: 'Outlook', requiresOptInConfirm: true, autoLaunch: false }` | +3 שורות |
| `outlook-wrapper.bat` | חדש, בתבנית של `word-wrapper.bat` | ~40 שורות |
| `scripts/tray-icon.ps1` | להוסיף ל-`$Apps`, Connect Outlook menu item, דיאלוג אזהרה (4.2), Disconnect Outlook only (4.6) | ~80 שורות |
| `scripts/inject.js` | filter קשיח (4.3), URL redaction בלוג (4.5), auto-disconnect timer per-app (4.4), detach per-app (4.6) | ~60 שורות |
| `scripts/port-discovery.js` | להסיר fallback `*.claude.ai` כשהאפליקציה Outlook | ~10 שורות |
| `doctor.bat` | להוסיף בדיקות: Outlook installed, Outlook process, Outlook port, Outlook target ב-apps.json | ~25 שורות |
| `install.bat` | אין שינוי - האפליקציות מתגלות אוטומטית |
| `docs/security.md` | סקציה חדשה: "Outlook-specific risks and mitigations" | ~50 שורות |

### מבנה state files - אין שינוי breaking

`claude-office-rtl.apps.json` יקבל מפתח רביעי: `{"Word":..., "Excel":..., "PowerPoint":..., "Outlook":...}`. גרסאות tray ישנות פשוט יתעלמו מהמפתח החדש.

`claude-word-rtl.status` נשאר aggregate על כל ארבעת האפליקציות.

---

## 6. שאלות פתוחות ל-probe (M0)

ענה על אלה לפני שמתחילים M1:

1. New Outlook (Monarch) vs Outlook הקלאסי - לתמוך באחד, בשניהם, או באחד-משניהם תחילה?
2. גרסאות Office: Microsoft 365 בלבד או גם Office 2021/2019 perpetual? (Claude add-in דורש 365.)
3. עברית ב-Outlook: ה-CSS הקיים מבוסס על MutationObserver על ה-DOM של pivot.claude.ai. אם ה-DOM זהה לחלוטין - עובד. אם Outlook host shell מכניס iframe נוסף מסביב - צריך התאמה.
4. האם להוסיף checkbox "השבת Outlook לחלוטין" ב-tray menu שמסתיר את כל הפיצ'ר עבור משתמשים שלא רוצים אותו? הצעה: כן, ולעשות אותו מסומן כברירת מחדל בהתקנה חדשה (opt-in גם ברמת ההפעלה הראשונה).

---

## 7. אבני דרך (Milestones)

| M | תוכן | הגדרת "DONE" |
|---|------|---------------|
| **M0** | probe על מכונה אמיתית, מענה ל-6 השאלות בסעיף 2 | `probe/outlook-host-discovery.js` ירוץ מקצה לקצה, דוח ב-`probe/README.md`, החלטת go/no-go. **בוצע 2026-05-12, GO לקלאסי, DEFER ל-New Outlook** |
| **M1** | הקשחות 4.1 (manualOnly, no auto-launch) + 4.3 (target filter קשיח ל-`_host_Info=Outlook$`) חייבות לקדום לקוד המינימלי - ראה סעיף 3 "עדכון מ-M0". אחריהן: `lib/office-apps.js` + `outlook-wrapper.bat` + פריט Connect Outlook ב-tray | filter ב-`inject.js` חוסם attach ל-Outlook אלא אם opt-in מפורש; manualOnly מונע auto-spawn של ה-injector על Outlook; Connect Outlook מציג את האייקון כירוק, CSS מוזרק, RTL נראה ב-panel |
| **M2** | יתר ההקשחות: 4.2 (דיאלוג אזהרה), 4.4 (auto-disconnect timer), 4.5 (URL redaction בלוג) | reviewer חיצוני (עוד claude code session או user) מאמת כל אחת מהשלוש |
| **M3** | Disconnect Outlook only (4.6), עדכון `doctor.bat`, עדכון `docs/security.md` | doctor.bat מציג 19 בדיקות (15 קיימות + 4 ל-Outlook), security.md כולל את הסעיף החדש |
| **M4** | תיעוד משתמש: עדכון README ו-README.he עם סקציית Outlook, CHANGELOG ל-v0.3.0, גירסה ב-package.json | קריאה של README ע"י מישהו שלא הכיר את הפרויקט מסבירה את הסיכון בצורה שאפשר להחליט |
| **M5** | release - smoke test על VM נקי (לפי `## Common commands` ב-CLAUDE.md, סעיף "Smoke test"), עדכון GitHub release | tag v0.3.0, release notes, install על מכונה שאינה של אסף |

---

## 8. תרחישי בדיקה ל-M5

לפני release, להריץ ידנית את כל אלה ולתעד תוצאות:

1. **Path הזהב:** Outlook סגור, install.bat נקי, tray ירוק, Connect Outlook, דיאלוג אזהרה, אישור, Outlook נפתח, פותח מייל בעברית, פותח Claude panel, מסכם מייל, RTL נראה תקין על תגובת Claude.
2. **Cancel בדיאלוג האזהרה:** Connect Outlook, Cancel - הסטטוס נשאר DISCONNECTED, Outlook לא נפתח, ה-tray לא משתנה.
3. **Concurrent:** Word + Excel + PowerPoint + Outlook כולם connected בו-זמנית. ארבעת ה-status labels ירוקים. הזרקת CSS עובדת בארבעתם.
4. **Auto-disconnect timer:** Connect Outlook, להמתין 15 דקות, לבדוק שהסטטוס עבר ל-DISCONNECTED ושהמשתמש צריך Connect מחדש.
5. **EDR:** להריץ על מכונה עם Defender ATP פעיל, לוודא שאין flags או blocks. אם יש - לתעד ולשקול חזרה.
6. **legacy users:** משתמש שמשדרג מ-v0.2.2 ל-v0.3.0 ולא מתעניין ב-Outlook - install.bat לא מציג שום שינוי משמעותי, ה-tray נראה אותו דבר עד שהמשתמש בעצמו מחפש Connect Outlook.
7. **Disconnect Outlook only:** Word ו-Outlook מחוברים, לחיצה על Disconnect Outlook only, Word נשאר ירוק, Outlook הופך אדום, Word panel ממשיך לפעול בלי הפרעה.

---

## 9. החלטות פתוחות לסשן הבא

- **התראה ויזואלית מתמשכת:** האם ה-tray יציג overlay אדום או badge כל זמן ש-Outlook מחובר, כדי שהמשתמש לא ישכח? הצעה: כן, "O" קטן בפינת האייקון.
- **גרסת default של auto-disconnect timer:** 15 דקות (סעיף 4.4) או 30, או 5?
- **שמירת היסטוריית CDP attaches:** ל-`%TEMP%\claude-word-rtl.log` להוסיף שורה לכל connect/disconnect של Outlook עם timestamp ובלי URL - לטובת audit אם המשתמש חושד בחדירה. ברירת מחדל: כן.
- **שם המוצר:** v0.2.x נקרא Claude for Office RTL Fix. עם Outlook זה עדיין "Office" טכנית. לא לשנות שם. שם הפרויקט נשאר.

---

## 10. סיכון ובאל-אאוט

אם M0 חוזר עם תשובה שלילית (Outlook לא ב-WebView2 / URL שונה / DOM שונה) - **ביטול התוכנית**, פרסום הממצא ב-CHANGELOG כ"considered, not pursued", והפניית משתמשי Outlook ל-Outlook on the Web עם browser extension אחר או workaround ידני.

אם M2 מגלה ש-EDR מסמן את הפעולה כ-suspicious (כמו v0.1.3) - **חזרה לתכנון**, אולי לחפש מנגנון injection שלא דרך CDP (למשל UI Automation, אבל זה כנראה לא ייתן את ה-CSS injection הנדרש).
