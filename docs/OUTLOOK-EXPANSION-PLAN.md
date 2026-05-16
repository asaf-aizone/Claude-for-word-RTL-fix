# Claude for Office RTL Fix - תוכנית הרחבה ל-Outlook

**סטטוס:** M0-M4 בוצעו, M5 (release smoke + GitHub release) פתוח. עדכון אחרון: 2026-05-16.
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
| **M1** | הקשחות 4.1 (manualOnly, no auto-launch) + 4.3 (target filter קשיח ל-`_host_Info=Outlook$`) חייבות לקדום לקוד המינימלי - ראה סעיף 3 "עדכון מ-M0". אחריהן: `lib/office-apps.js` + `outlook-wrapper.bat` + פריט Connect Outlook ב-tray | filter ב-`inject.js` חוסם attach ל-Outlook אלא אם opt-in מפורש; manualOnly מונע auto-spawn של ה-injector על Outlook; Connect Outlook מציג את האייקון כירוק, CSS מוזרק, RTL נראה ב-panel. **בוצע 2026-05-13/14 בארבעה sub-commits: M1a `b324bd7` (opt-in gate ב-inject.js: BLOCKED_HOST_INFO_KEYS + OPTIN_FLAGS), M1b `f4f5c9c` (רישום Outlook ב-office-apps APPS + tray $Apps), M1c `847182d` (outlook-wrapper.bat), M1d `fe23b31` (Connect Outlook tray menu + דיאלוג אזהרה per-launch - מכסה גם את 4.2 מתוך M2)** |
| **M2** | יתר ההקשחות: 4.2 (דיאלוג אזהרה), 4.4 (auto-disconnect timer), 4.5 (URL redaction בלוג) | reviewer חיצוני (עוד claude code session או user) מאמת כל אחת מהשלוש. **בוצע 2026-05-14: 4.2 בוצע כחלק מ-M1d `fe23b31`; 4.4 ו-4.5 בוצעו ב-`0cefc72` (URL redaction בלוג + 15-minute auto-disconnect timer ב-inject.js)** |
| **M3** | Disconnect Outlook only (4.6), עדכון `doctor.bat`, עדכון `docs/security.md` | doctor.bat מציג 19 בדיקות (15 קיימות + 4 ל-Outlook), security.md כולל את הסעיף החדש. **בוצע 2026-05-16: `a992284` (Disconnect Outlook only IPC ב-inject.js + tray, doctor 19 בדיקות, סעיף Outlook ב-security.md) + `ecb71f4` (early race guard ב-attach() אחרי post-fix review שגילה שה-late guard מנקה רק CDP אחרי DOM mutation)** |
| **M4** | תיעוד משתמש: עדכון README ו-README.he עם סקציית Outlook, CHANGELOG ל-v0.3.0, גירסה ב-package.json | קריאה של README ע"י מישהו שלא הכיר את הפרויקט מסבירה את הסיכון בצורה שאפשר להחליט. **בוצע 2026-05-16 ב-`b5b21df`: package.json 0.2.2->0.3.0, CHANGELOG [0.3.0] עם כל commits M0-M3 + הפניות לקוד, README.md (חלק עברי + אנגלי) עם סקציה ייעודית "Outlook (opt-in)", שלוש שורות חדשות בטבלת תפריט ה-tray, שורה חדשה בטבלת access, callout באבטחה, ו-doctor 15->19. README.he.md עם עדכונים מקבילים בקטגוריות תמציתיות.** |
| **M5** | release - smoke test על VM נקי (לפי `## Common commands` ב-CLAUDE.md, סעיף "Smoke test"), עדכון GitHub release | tag v0.3.0, release notes, install על מכונה שאינה של אסף |

---

## 7a. תיקונים שלא תוכננו במקור

2603283 (2026-05-14) - תיקון race ב-Connect Word/Excel/PowerPoint שהתגלה אחרי M2. ה-mutex של single-instance של Office חי 1-3 שניות אחרי ש-Get-Process מחזיר ריק, וגרם ל-launch השני לצאת בשקט. תוקן ע"י polling של tasklist ב-wrappers (5 איטרציות x 1s) + הארכת ה-delay ב-tray מ-500ms ל-1500ms + תיקון רגרסיה במסלול force-close. לא נוגע ב-Outlook (ה-wrapper שלו כבר מסרב לרוץ אם Outlook חי).

---

## 8. תרחישי בדיקה ל-M5

לפני release, להריץ ידנית את כל אלה על מכונת בדיקה (לא ה-dev machine) ולתעד תוצאות ב-issue/PR description. כל תרחיש מציין observation criteria מדויקים - לא "נראה תקין" אלא "ראית שורה X בלוג Y" / "כפתור Z עבר ל-Enabled תוך N שניות".

### 8.0 - הכנות לפני שמתחילים

**Tooling שצריך:**
- PowerShell עם `Get-Process`, `Get-CimInstance` (מובנה ב-Windows 10/11).
- `netstat -ano` (מובנה).
- `Get-Content -Wait -Path %TEMP%\claude-word-rtl.log` ב-PowerShell נפרד כדי לעקוב אחר הלוג בזמן אמת.
- DevTools attach: פתיחת Edge בכתובת `edge://inspect`, או `chrome://inspect` ב-Chromium, ו-Configure → Add `localhost:<port>` כאשר ה-port מגיע מ-`netstat -ano | findstr msedgewebview2-PID`.
- בדיקת EDR: `Get-MpThreatDetection` (Defender PowerShell module). לא קריטי אם המכונה ללא ATP - דלגו על תרחיש 5 ותעדו "no EDR available".

**סדר ההרצה:**
- תרחיש 6 (legacy upgrade) **חייב** לרוץ ראשון על snapshot נפרד של VM עם v0.2.2 מותקן מראש. אחרת ה-baseline ל-comparison חסר.
- תרחישים 1-5 ו-7 על snapshot נקי עם v0.3.0 fresh install.
- תרחישים 8-17 על אותו snapshot של 1-5 ו-7 או על snapshot חדש - לא חשוב.

**ניקוי בין תרחישים:** אחרי כל תרחיש שמשאיר state (1, 3, 4, 7, 8, 9, 10, 15, 17 - כל מה שעושה Connect): right-click tray → **Disconnect all**, ואז `del %TEMP%\claude-office-rtl.*` ב-cmd. אחרי תרחיש 4 (אם הופחת `OUTLOOK_AUTO_DISCONNECT_MIN`): `git checkout scripts/inject.js` ולהפעיל `git diff scripts/inject.js` כדי לוודא חזרה.

**Time budget:** באומדן הוגן, ~6-9 שעות לטסטר בודד לכל 17 התרחישים (תרחיש 4 לבדו ~15-17 דקות המתנה אם לא מפחיתים את הטיימר; 9 ו-10 כוללים מחזורי relaunch של Outlook; 5 דורש imaging של EDR).

### 8.1 - 7 תרחישי ה-baseline

1. **Path הזהב.**
   - Pre-req: Outlook הקלאסי סגור, `olk.exe` סגור (חובה - ה-wrapper יסרב אם הוא רץ), install.bat רץ נקי על המכונה.
   - צעדים: ממתינים עד 30 שניות שאייקון ה-tray יעבור מאפור לאדום (ה-injector טרם עלה כי אין הסכמה ל-Outlook). Right-click → Connect Outlook → דיאלוג עם Cancel ממוקד (לראות מסגרת מיקוד סביב הכפתור Cancel) → לחיצה על OK → Outlook נפתח. **טווח הזמן תלוי בנתיב:** הפעלה ראשונה (ה-injector לא רץ עדיין) ~3-5 שניות בגלל `timeout /t 3` ב-`outlook-wrapper.bat:78`. אם ה-injector כבר רץ מ-Connect קודם של Word/Excel/PowerPoint - פחות משנייה.
   - Pass: בתוך 5-10 שניות מהפתיחה, ה-status label של Outlook עובר ל-`connected`, האייקון ירוק. לוג ב-`%TEMP%\claude-word-rtl.log` מציג שורה שמתחילה ב-`Attached & injected: [Outlook]`. פתיחת מייל בעברית, לחיצה על "Summarize this email" - תגובת Claude נראית RTL (טקסט מיושר ימינה, scrollbar בשמאל, סימני רשימה בצד ימין).

2. **Cancel בדיאלוג האזהרה.**
   - Pre-req: Outlook סגור.
   - Sub-test 2a (לחיצה על Cancel): Connect Outlook → Cancel → אישור ש: לא נכתב opt-in flag ב-`%TEMP%\claude-office-rtl.outlook-optin`, Outlook לא נפתח, ה-tray לא משתנה.
   - Sub-test 2b (Enter ב-default-Cancel): Connect Outlook → לחיצה על Enter בלי לזוז עם העכבר → התוצאה זהה ל-Cancel. זה ה-regression test לעיצוב `MessageBoxDefaultButton::Button2` ב-`tray-icon.ps1:456`.

3. **Concurrent: 4 אפליקציות מחוברות בו-זמנית.**
   - Pre-req: כל ארבע סגורות.
   - צעדים: Connect Word → ירוק → Connect Excel → ירוק → Connect PowerPoint → ירוק → Connect Outlook (דיאלוג, OK) → ירוק.
   - Pass: ארבע ה-status labels = `connected`. **חובה לקרוא את** `%TEMP%\claude-office-rtl.apps.json` ולוודא: `{"Word":"CONNECTED","Excel":"CONNECTED","PowerPoint":"CONNECTED","Outlook":"CONNECTED"}`. הזרקת CSS עובדת בכל ארבעתם (פתיחת פאנל Claude בכל אחד והבחנה ב-RTL).

4. **Auto-disconnect timer.** **[DEV-ONLY אם משנים את הטיימר]**
   - Pre-req: `OUTLOOK_AUTO_DISCONNECT_MIN = 15` ב-`scripts/inject.js:82` בקוד הנשלח. **אופציה לבדיקה מעשית:** לערוך זמנית ל-`OUTLOOK_AUTO_DISCONNECT_MIN = 1` ב-`scripts/inject.js:82` (שורה יחידה, מספר אחד). **גארד הכרחי לפני tag:** להריץ `git diff scripts/inject.js` ולוודא שהפלט ריק לחלוטין; הוספת CI grep guard ל-`OUTLOOK_AUTO_DISCONNECT_MIN = [0-9]+` שמסמן fail אם הערך לא בדיוק 15 היא רעיון לעתיד.
   - צעדים: לעצור את ה-injector (right-click tray → Disconnect all), לערוך את הקובץ, לשמור, להפעיל Connect Outlook מחדש. להשאיר את הפאנל פתוח. הטיימר מתחיל בזמן ה-`Page.addScriptToEvaluateOnNewDocument` ב-attach (`inject.js:494`), לא בזמן פעילות של המשתמש. להמתין 15 דקות (או 1 דקה).
   - Pass: בלוג מופיע שורה שמכילה `Outlook auto-disconnect timeout (15 min)` (או `(1 min)` אם הופחת). ה-status label של Outlook עובר ל-`DISCONNECTED`. הקובץ `%TEMP%\claude-office-rtl.outlook-optin` נעלם (בדיקה: `dir %TEMP%\claude-office-rtl.outlook-optin` מחזיר File Not Found). Connect Outlook נדרש שוב כדי לחדש (יחזיר דיאלוג).

5. **EDR (Defender ATP / equivalent).** **[EDR-LAB-ONLY - לדלג אם לא זמין, לתעד "no EDR available"]**
   - Pre-req: מכונת בדיקה עם Microsoft Defender ATP / CrowdStrike / SentinelOne פעיל.
   - צעדים: install.bat, Connect Word, Connect Outlook.
   - Pass (**כל ארבעת התנאים נדרשים**, AND לא OR):
     - (a) אין notification toast מ-Defender / EDR האחר בזמן ההפעלה.
     - (b) `Get-MpThreatDetection` (Defender PowerShell module) לא מחזיר entry שמצביע על `inject.js` או על ה-wrappers.
     - (c) ה-wrappers מסיימים עם exit code 0. בדיקה: בכל cmd window שהפעיל wrapper, `echo %ERRORLEVEL%` מחזיר 0.
     - (d) `msedgewebview2.exe` שומר את ה-LISTENING port. **איך למצוא את ה-PID הנכון:** `Get-Process msedgewebview2 | Where-Object { $_.MainWindowTitle -match 'Outlook' -or $_.Parent.Name -eq 'OUTLOOK' }` ואז `netstat -ano | findstr <PID>` - מחפשים שורת `LISTENING` עם הפורט.
   - Fail action: לתעד את ה-flag/block ולשקול חזרה (אופציה: rollback ל-v0.2.2 קוד החיבור של Outlook).

6. **Legacy users (משדרגים שלא מתעניינים ב-Outlook).** **[חייב לרוץ ראשון על snapshot עם v0.2.2 מותקן מראש]**
   - Pre-req: v0.2.2 מותקן ופעיל על snapshot של VM. **Step 0:** לפני ה-upgrade, לרשום את ספירת פריטי תפריט ה-tray הקיים (לקחת screenshot של התפריט הפתוח).
   - צעדים: להריץ install.bat של v0.3.0 על אותה תיקייה.
   - Pass:
     - (a) ספירת פריטי תפריט גדלה ב-3 בדיוק לעומת ה-screenshot: Outlook status label (מנוטרל), Connect Outlook, Disconnect Outlook only. שאר הפריטים זהים.
     - (b) אין דיאלוג אזהרה ב-startup של ה-tray.
     - (c) הקובץ `%TEMP%\claude-office-rtl.outlook-optin` לא קיים עד שהמשתמש לוחץ Connect Outlook באופן יזום.
     - (d) `doctor.bat` מציג 19 בדיקות (היו 15) אבל 16-19 הן `:info` גם במכונה ללא Outlook מותקן.
     - (e) אם Outlook פתוח בזמן הבדיקה, הלוג מציג שורה אחת `Blocked target (no opt-in): outlook` בפעם הראשונה שה-target נצפה, ולא חוזרת בכל tick (קוד `loggedBlockedIds` ב-`inject.js` מבטיח de-duplication עד שה-target נעלם). אם Outlook נסגר ונפתח שוב - השורה תופיע שוב פעם אחת.

7. **Disconnect Outlook only.**
   - Pre-req: Outlook מחובר (תרחיש 1) + לפחות אחת מ-Word/Excel/PowerPoint מחוברת.
   - צעדים: לחיצה על Disconnect Outlook only.
   - Pass: (a) הקובץ `%TEMP%\claude-office-rtl.disconnect-outlook.request` מופיע ונעלם בתוך ~2 שניות (קצב ה-tick), (b) בלוג: `Disconnect-only request honoured for [Outlook]: closedAny=true`, (c) Word (או האפליקציה האחרת) **לא** מקבלת שורת `Detached` בלוג, (d) ה-status של Outlook עובר ל-`DISCONNECTED` אבל של Word נשאר `connected`, (e) `%TEMP%\claude-office-rtl.outlook-optin` נעלם.

### 8.2 - 10 תרחישי regression ייעודיים ל-v0.3.0

תרחישים שתופסים באגים שתועדו במהלך הפיתוח של M0-M4.

8. **Race condition בזמן Connect Outlook handshake (regression ל-`ecb71f4`).** **[DEV-ONLY - reaction time אנושית לא אמינה ב-≤500ms]**
   - הסבר: ה-race שתועדה היא קליק על Disconnect Outlook only בזמן ה-3 awaits של ה-CDP handshake (CDP enable → Page.enable → Runtime.enable). זה חלון של מילישניות בודדות עד עשרות.
   - אופציות בדיקה:
     - **a (אמינה - dev):** להוסיף `setTimeout(() => {}, 2000)` זמני בין `await Page.enable()` ו-`await Runtime.enable()` ב-`inject.js`, אז Connect → המתן 1s → Disconnect Outlook only ידני. **למחוק את ה-setTimeout** לפני tag.
     - **b (לא אמינה - אדם):** שני אנשים, אחד לוחץ Connect Outlook + OK, השני לוחץ Disconnect Outlook only מיד אחרי הקליק על OK. אם נכשל, מספיק תיעוד "race window not reproducible by hand - dev sub-agent verified at commit ecb71f4".
   - Pass: בלוג מופיע אחת משתיים: שורה שמכילה `Attach aborted pre-evaluate: [Outlook]` (אם נתפס בשלב המוקדם) או `Attach aborted post-handshake: [Outlook]` (אם נתפס מאוחר). ה-status של Outlook נשאר `DISCONNECTED`. ה-opt-in flag נעלם.

9. **Connect Outlook כש-Outlook כבר רץ (זרם שונה ב-`Start-ConnectOutlook`).**
   - Pre-req: Outlook פתוח עם טיוטה (לא חובה לא-שמורה - האחריות על draft recovery היא של Outlook עצמו, לא של הכלי).
   - צעדים: Connect Outlook → OK בדיאלוג הראשון (warning) → OK בדיאלוג השני ("close and relaunch") → ממתינים עד שOutlook נסגר בנימוס.
   - Hard timeout: 10 שניות (`tray-icon.ps1:523` - `WaitedMs -ge 10000`). אם Outlook לא נסגר תוך 10 שניות - תרחיש 10 מטפל בזה (force-close). כאן אנחנו בודקים את ה-graceful path בלבד.
   - Pass: Outlook נסגר תוך ≤10 שניות, נפתח מחדש דרך ה-wrapper, status flips ל-`connected`, בלוג: שורה שמתחילה ב-`Attached & injected: [Outlook]`. **הערה:** התנהגות AutoSave של Outlook (האם טיוטה משוחזרת) היא של Outlook ולא של הכלי הזה - לא לבדוק.
   - Sub-test 9a: Cancel בדיאלוג השני - Outlook לא נסגר, status נשאר `running without RTL`, ה-opt-in flag לא נכתב.

10. **Force-close path (regression ל-`tray-icon.ps1:527-556`).**
    - Pre-req: Outlook פתוח עם דיאלוג modal פתוח שחוסם shutdown. **איך ליצור:** לפתוח Compose mail חדש, להקליד תוכן, ללחוץ X של חלון Outlook (לא Send) - מופיע prompt "Save changes? / Send / Discard". להשאיר אותו פתוח.
    - צעדים: Connect Outlook → OK בדיאלוג ה-warning → OK בדיאלוג ה-"close and relaunch" → ממתינים 10 שניות → דיאלוג "did not close within 10 seconds" → OK → force-kill → relaunch דרך wrapper.
    - Pass: אחרי force-kill, Outlook נפתח מחדש דרך wrapper. data loss של מה שלא נשמר ב-modal dialog צפוי וזה תקין (האזהרה בדיאלוג ה-force-close מסבירה את זה).
    - Sub-test 10a: Cancel בדיאלוג של ה-10 שניות - Outlook נשאר במצבו (עם ה-modal פתוח), status נשאר `running without RTL`, אין force-kill.

11. **Outlook ללא תוסף Claude מותקן.**
    - Pre-req: Outlook הקלאסי מותקן אבל הוסר/הושבת ה-Claude add-in.
    - צעדים: Connect Outlook → OK → Outlook נפתח דרך ה-wrapper.
    - Pass: status נשאר `running without RTL` (לא עובר ל-`connected`). בלוג **אין** שורת `Attached & injected: [Outlook]`. הלוג כן מציג `targets` ללא Claude target עבור Outlook. ה-tray לא מציג שום שגיאה - פשוט אין חיבור כי אין מה לחבר.

12. **Two tray instances (mutex regression).**
    - Sub-test 12a: `wscript scripts\start-tray.vbs` פעמיים → instance שני יוצא מיידית, רק אייקון אחד באזור ההודעות. מאשרים עם `Get-Process powershell` שיש רק powershell.exe אחד עם command line שמכיל `tray-icon.ps1`.
    - Sub-test 12b (recovery after force-kill): kill ל-tray הראשון ב-`Stop-Process -Force` → הפעלת `start-tray.vbs` שוב. **התנהגות צפויה ומדויקת:** הקוד ב-`tray-icon.ps1:42` משתמש ב-`New-Object Mutex($true, name, [ref]$createdNew)` ויוצא מיידית אם `-not $createdNew`. כש-owner נהרג ב-force-kill, kernel object של ה-mutex נשאר עד שכל handles אליו נסגרים (rundown של תהליך) - בזמן הזה, ה-tray השני יראה `$createdNew=false` ויצא בשקט. **ה-recovery לא מיידי**: יש לחכות ש-Windows יסיים את ה-rundown של התהליך הראשון (בד"כ שניות בודדות) ואז `start-tray.vbs` יצליח. Pass = אחרי המתנה ≥3 שניות, ה-tray השלישי שמופעל מצליח לעלות. אם נכשל אחרי 10 שניות - escalation, do not retry.

13. **Corrupted opt-in flag file.**
    - Sub-test 13a: ידנית `Set-Content -Path $env:TEMP\claude-office-rtl.outlook-optin -Value '' -NoNewline -Encoding ascii` (קובץ ריק בלי BOM/newline) → Connect Outlook → ה-injector מתחבר. זה התנהגות צפויה - הבדיקה ב-`inject.js:449` היא `fs.existsSync` בלבד, ללא בדיקת תוכן או גודל. **הקיום לבד הוא הסכמה.**
    - Sub-test 13b (security observation - **INFORMATIONAL, לא חוסם release**): בלי לחיצה על Connect Outlook, ידנית לכתוב את ה-flag ולפתוח Outlook ידנית דרך `outlook-wrapper.bat` - ה-injector מתחבר. **זה bypass של הדיאלוג ו-known limitation:** משתמש או תוכנה שכבר רצים תחת אותו user יכולים לכתוב את ה-flag ולעקוף את ההסכמה הוויזואלית. ההגנה היא לא בקובץ עצמו אלא בעובדה שתוקף שכבר יכול לכתוב ל-`%TEMP%` כבר יכול לקרוא את ה-CDP ישירות (אותו threat boundary). מטופל בעתיד אם נדרש (לא בהיקף v0.3.0): provenance check שה-flag נכתב מתהליך wrapper לגיטימי.

14. **olk.exe (New Outlook) פתוח.**
    - Pre-req: New Outlook (`olk.exe`) פתוח, Outlook הקלאסי סגור.
    - צעדים: Connect Outlook.
    - Pass: מופיע דיאלוג "New Outlook is running, close it first" (`tray-icon.ps1:421-432`). ה-wrapper לא רץ. ה-opt-in flag לא נכתב. status נשאר אותו דבר.

15. **Wrapper run ישירות (bypassing tray).**
    - צעדים: double-click ידני על `outlook-wrapper.bat` מ-File Explorer (entry point נתמך).
    - Pass: ה-wrapper כותב את ה-opt-in flag, מפעיל את ה-injector אם הוא לא רץ, פותח את Outlook. ה-tray (אם רץ) מזהה את המצב תוך 2 שניות ומציג את Outlook = `connected`.

16. **Disconnect Outlook only כש-Outlook לא מחובר (idempotency).**
    - Pre-req: Outlook לא מחובר (`disconnect` או `not running`).
    - צעדים: לחיצה על Disconnect Outlook only.
    - Pass: הפריט עצמו amור להיות disabled (לפי `tray-icon.ps1:1056` - מותר רק כש-Outlook = `connected`). אם איכשהו לוחצים עליו (למשל timing race לפני tick): הקובץ נכתב, ה-injector צורך אותו, בלוג `closedAny=false`, שום שגיאה ל-user.

17. **Auto-disconnect timer survives navigation.**
    - Pre-req: Connect Outlook לפני <14 דקות. אופציה לבדיקה מעשית: להפחית `OUTLOOK_AUTO_DISCONNECT_MIN` ל-2 (גארד `git diff` כמו בתרחיש 4).
    - צעדים: ניווט בין folders של מייל ב-Outlook (סדר Inbox → Sent → Drafts). פתיחת compose window חדש. כל אחת מהפעולות האלה מפעילה `frameNavigated` ב-CDP על ה-WebView2 host של תוסף Claude.
    - Pass: הטיימר ממשיך לרוץ מזמן ה-`Page.addScriptToEvaluateOnNewDocument` המקורי. ה-`frameNavigated` handler ב-`inject.js:515-522` מריץ את `INJECTOR_SCRIPT` שוב (לוודא ב-DevTools שה-`<style id="__claude_rtl_fix__">` עדיין קיים), אבל **לא** קורא ל-`clearTimeout` ו-**לא** מאתחל setTimeout חדש. כלומר אחרי 15 דקות (או 2 אם הופחת) מה-attach המקורי - disconnect, ללא קשר לכמה navigations היו בדרך.
    - איך לאשר: open DevTools על ה-msedgewebview2 PID של Outlook (`edge://inspect` → Configure → `localhost:<port>`), inspect את ה-Claude panel. אחרי הניווטים, אמור עדיין להיות `<style id="__claude_rtl_fix__">` ב-document.head.

### הערות לתיעוד תוצאות

- כל תרחיש שנכשל - לתעד את הגרסה המדויקת (commit SHA), המכונה, גרסת Windows ו-Office, ולהדביק את `%TEMP%\claude-word-rtl.log` ואת `doctor.log`.
- תרחישים 1-7 הם תנאי הכרחי לתג. תרחישים 8-17 הם מומלצים בחום אבל ניתן לדחות אם זה גוזל יותר מ-4 שעות בדיקה ולתעד אותם בפרטים ב-issue פתוח שייסגר במהלך v0.3.1.
- תרחישים מסומנים `[DEV-ONLY]` (4 ו-8 עם reduction של הטיימר/setTimeout) או `[EDR-LAB-ONLY]` (5) - אם אין סביבה מתאימה, לדלג ולתעד את הסיבה. לא חוסם release.
- Sub-test 13b מסומן `INFORMATIONAL` - הוא תיעוד של known limitation, לא test של פיצ'ר. אין לרשום אותו כ-fail אם ה-bypass עובד; זו ההתנהגות הצפויה.

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
