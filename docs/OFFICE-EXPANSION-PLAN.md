# Claude for Office RTL Fix - תוכנית הרחבה לכל אפליקציות Office

**סטטוס:** טיוטה לביצוע ביום אחר
**גרסת יעד:** v0.2.0
**מסמך מקור:** [probe/README.md](../probe/README.md), [CHANGELOG.md](../CHANGELOG.md)

---

## 1. החזון

אייקון אחד על שולחן העבודה ובתפריט התחל שנקרא **"Claude for Office RTL"**. המשתמש לוחץ עליו > רואה חלון קטן עם שלושה כפתורים (Word, Excel, PowerPoint) > בוחר אפליקציה > היא נפתחת עם תיקון RTL פעיל ב-Claude add-in. ברקע, injector יחיד מזהה ומטפל בכל שלוש האפליקציות אוטומטית.

בנוסף: double-click על קובץ docx/xlsx/pptx יפעיל את התיקון אוטומטית (ע"י file associations).

אייקון מגש (tray) מציג סטטוס מצטבר של שלוש האפליקציות.

---

## 2. ממצאי שלב 1 (probe)

בוצע 2026-04-22. שתי האפליקציות נבדקו מול WebView2 debug port:

| אפליקציה | Port | URL | ה-pattern הקיים תופס? |
|----------|------|-----|-----------------------|
| Word | 9222 | `https://pivot.claude.ai/?...&_host_Info=Word$Win32$16.01$he-IL` | כן (primary) |
| Excel | 9223 | `https://pivot.claude.ai/?...&_host_Info=Excel$Win32$16.01$he-IL` | כן (primary) |
| PowerPoint | 9224 | `https://pivot.claude.ai/?...&_host_Info=Powerpoint$Win32$16.01$he-IL` | כן (primary) |

**מסקנה:** שלושתן משתמשות באותו URL של pivot.claude.ai, אותו DOM, אותה UI shell. לוגיקת ה-CSS וה-MutationObserver תעבוד כמות שהיא. אין צורך בשינוי ב-`INJECTOR_SCRIPT` או ב-`RTL_CSS`.

---

## 3. הבעיה הארכיטקטונית המרכזית: ה-auto-enable

### מצב קיים (v0.1.0)

התפריט של tray-icon.ps1 מאפשר "Auto-Enable" שכותב:
```
HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=9222"
```

זה גורם לכל Word שנפתח (מ-Recent, מ-taskbar, מ-double-click) לחשוף port 9222 אוטומטית.

### למה זה נשבר עם 3 אפליקציות

env var זו **מערכתית למשתמש**, אחת בלבד. אם הערך הוא `--remote-debugging-port=9222` אז:
- Word נפתח > תופס 9222 > עובד
- Excel נפתח **אחריו** > מנסה לתפוס 9222 > **נכשל שקט** (WebView2 לא נופל, פשוט לא פותח debug)
- PowerPoint נפתח **אחריו** > אותו כישלון שקט

רק האפליקציה הראשונה שנפתחה תיהנה מ-RTL. זו regression חמורה מול v0.1.0.

### החלופות

| חלופה | יתרון | חיסרון |
|--------|--------|---------|
| **A. ביטול auto-enable לגמרי, wrappers בלבד** | דטרמיניסטי, תומך ב-3 אפליקציות בו-זמנית | double-click docx מ-Explorer לא יעבוד בלי file associations |
| **B. `--remote-debugging-port=0` (דינמי)** | כל אפליקציה מקבלת port אוטומטית | גילוי ה-port מורכב (צריך לחפש ב-DevToolsActivePort file או לסרוק ports), שביר |
| **C. auto-enable רק ל-Word, wrappers ל-Excel/PowerPoint** | תאימות לאחור מלאה | התנהגות לא עקבית, מבלבל |
| **D. wrappers + file associations + AppPaths** | פועל בכל דרך שהמשתמש פותח בה Office | הכי הרבה קוד, צריך להסתייג משיוכי קבצים קיימים |

**הבחירה הנוכחית (2026-04-25): חלופה B - dynamic ports + AUTO.** ראה סעיף 3.5 למטה.

המלצה מקורית הייתה חלופה D, אבל אחרי דיון נוסף הוחלט לחזור לחלופה B. הסיבה: דרך `--remote-debugging-port=0` כל אפליקציה תופסת פורט פנוי בעצמה, ה-injector מגלה אותו דרך `tasklist`+`netstat`, וה-UX הופך ל"פעם אחת מפעילים Auto-Enable וזהו" - בלי wrappers, בלי file associations, בלי launcher כדרך ראשית.

### 3.4 ממצאי M0 - POC validation (2026-04-25)

POC הורץ בהצלחה על המכונה של אסף. שלושת אפליקציות Office (Word, Excel, PowerPoint) הופעלו בו-זמנית עם `--remote-debugging-port=0`, כל אחת עם Claude add-in פתוח. תוצאות:

| App | WebView2 PID | Port דינמי |
|-----|-------------|-----------|
| Excel | 12896 | 55099 |
| Word | 31492 | 61918 |
| PowerPoint | 31492 | 61918 |

**ממצא מפתיע - Office מאחד WebView2 hosts:** Word ו-PowerPoint **חולקים את אותו תהליך WebView2** (PID 31492, port 61918). אותו תהליך מארח 2 page targets - אחד עם `_host_Info=Word$...` ואחד עם `_host_Info=Powerpoint$...`. Excel קיבל תהליך משלו (כנראה כי הוא נפתח ראשון).

**השלכות על המימוש:**
1. אסור להניח "פורט אחד לאפליקציה". ההנחה הנכונה: **1 עד 3 פורטים דינמיים, כל פורט מארח 1 עד 3 Claude targets**.
2. זיהוי האפליקציה נעשה אך ורק לפי `_host_Info=` בתוך URL ה-target, לא לפי פורט.
3. הלוגיקה הזו בעצם כבר קיימת ב-`inject.js` הנוכחי - הוא מבצע `listTargets` ועוצר על Claude target. רק צריך להרחיב לגלות אילו פורטים בכלל קיימים.
4. נחסך עומס: במקום 3 חיבורי WebSocket לכל ענייני התדפיסים, יכול להיות 1-2 חיבורים בלבד.

**ממצא נוסף:** ב-Auto-Enable הקיים (HKCU `port=9222`), זוהו 2 תהליכי WebView2 של תוכנות צד שלישי (LogiAiPromptBuilder, TeamViewer) שירשו את ה-env var ותפסו את 9222 לפני Office. זה מחזק את ההצדקה למעבר ל-`port=0` - מנע התנגשויות שקטות עם תוכנות אחרות.

**POC נחתם כתקין.** אפשר לעבור ל-M1.

---

### 3.5 הארכיטקטורה החדשה - Dynamic Ports + AUTO

**Auto-Enable env var (מעודכן):**
```
HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=0"
```

`--remote-debugging-port=0` אומר ל-WebView2: "בחר פורט פנוי בעצמך". כל הפעלה של Word/Excel/PowerPoint תופסת פורט אקראי בטווח dynamic ports של Windows (49152-65535) ולא מתנגשת.

**מנגנון גילוי פורטים ב-injector:**

```
1. tasklist /FI "IMAGENAME eq msedgewebview2.exe" /FO CSV
   > רשימת PIDs של תהליכי WebView2 פעילים

2. netstat -ano | findstr LISTENING
   > מיפוי PID > פורט שמאזין

3. בכל פורט מועמד: GET http://localhost:<port>/json/list

4. סינון לפי URL_PATTERN_PRIMARY (pivot.claude.ai)
   זיהוי האפליקציה לפי _host_Info= בתוך ה-URL
```

**יתרונות מול חלופה D:**
- אין צורך ב-3 wrappers נפרדים (`excel-wrapper.bat`, `powerpoint-wrapper.bat`)
- אין צורך ב-file associations חיוניים
- Launcher UI הופך לאופציונלי (nice to have)
- Auto-Enable **נשאר** - רק עם ערך מעודכן
- חוויית משתמש: לחיצה אחת ב-install, ומכאן והלאה כל פתיחה של Office מכל מקור (taskbar, double-click, recent files, OneDrive) פועלת אוטומטית

**עלות:**
- POC קצר ב-`probe/` כדי לוודא ש-`tasklist`+`netstat` מחזירים מה שצריך על מערכת אמיתית
- לוגיקת port discovery חדשה ב-`inject.js` במקום `const PORT = 9222`
- ה-tray ידע לדווח על "X מתוך 3 אפליקציות מחוברות" לפי מספר ה-targets שזוהו

---

## 4. ארכיטקטורת היעד

```
  Auto-Enable (חד-פעמי בהתקנה):
  HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0

         המשתמש פותח Office בכל דרך:
   ┌──────────────┬─────────────┬──────────────┬─────────────┐
   │ Taskbar      │ Double-click│ Recent files │ OneDrive,   │
   │              │ (.docx etc) │              │ email, etc. │
   └──────┬───────┴──────┬──────┴──────┬───────┴──────┬──────┘
          ▼              ▼             ▼              ▼
     WINWORD.EXE    EXCEL.EXE    POWERPNT.EXE   (any combination)
          │              │             │
          ▼              ▼             ▼
    [WebView2 picks   [WebView2 picks  [WebView2 picks
     free port X]      free port Y]    free port Z]
          │              │             │
          └──────────────┼─────────────┘
                         ▼
       ┌──────────────────────────────────────┐
       │   inject.js (1 instance)              │
       │   - tasklist: msedgewebview2.exe PIDs │
       │   - netstat: PID > LISTENING port     │
       │   - probe each port for Claude target │
       │   - identify app via _host_Info=      │
       └──────────────────┬───────────────────┘
                          ▼
       ┌──────────────────────────────────────┐
       │   tray-icon.ps1                       │
       │   status: "X of N apps connected"     │
       └──────────────────────────────────────┘

   Optional helpers (not required for AUTO to work):
   - Launcher UI (WPF/WinForms) - explicit "open Office with RTL" button
   - File associations - opt-in at install time
   - Per-app wrappers - command-line entry for power users / scripts
```

**עקרונות:**
- **Auto-Enable מערכתי הוא הנתיב הראשי.** משתמש מפעיל פעם אחת > כל פתיחה של Office עובדת אוטומטית.
- **Port discovery דינמי** ב-injector - אין רשימת ports קבועה. ה-injector מגלה אילו תהליכי `msedgewebview2.exe` חיים, איזה פורט הם תופסים, ובוחן אם יש Claude target.
- **אפליקציה זוהית לפי `_host_Info=`** ב-URL של ה-target, לא לפי הפורט.
- injector יחיד. lock + pid אחד. ידידותי לפתיחות מרובות.
- tray status: "X של N אפליקציות פתוחות מחוברות". אם אין אף Office פעיל > אפור. אם יש Claude target בכל ה-WebView2 הפעילים > ירוק. שגיאת port discovery > אדום.
- Launcher UI, file associations, per-app wrappers - **כולם אופציונליים ולא חיוניים** למסלול ה-AUTO.

---

## 5. אבני דרך

### M0 - POC validation (חוסם הכול)
לפני כל קוד production - אימות שהארכיטקטורה החדשה עובדת על מערכת אמיתית.

**קבצים (כבר קיימים ב-`probe/`):**
- `probe/launch-office-dynamic.bat` - מפעיל Word/Excel/PowerPoint עם `--remote-debugging-port=0`
- `probe/dynamic-port-discovery.js` - מגלה PIDs+ports דרך tasklist+netstat, מזהה אפליקציה לפי `_host_Info`

**קריטריוני הצלחה:**
- 3 אפליקציות פתוחות בו-זמנית > 3 PIDs של `msedgewebview2.exe` שמאזינים על 3 פורטים שונים בטווח 49152-65535.
- כל פורט תופס Claude target עם `_host_Info=Word/Excel/Powerpoint`.
- ה-script מציג `SUCCESS: dynamic-port architecture works`.

**אם נכשל:** חוזרים לחלופה D המקורית (3 wrappers + file associations). מתעדים את הסיבה ב-`docs/bugs/bug-NNN-dynamic-port-poc-failed.md`.

### M1 - Port discovery ב-injector + Auto-Enable מעודכן (ליבה)
המעבר של ה-MVP מקוד POC ל-production.

**קבצים חדשים:**
- `lib/office-apps.js` - metadata קל: `[{ name, processName, urlHostInfo }, ...]` ללא ports (כי הם דינמיים).
- `scripts/port-discovery.js` - מודול שעוטף את הלוגיקה של tasklist+netstat+probe (חילוץ מ-`probe/dynamic-port-discovery.js` עם הקשחה לprodution).

**שינויים ב-[scripts/inject.js](../scripts/inject.js):**
- הסרת `const PORT = 9222`. במקום `listTargets(port)` קבוע > קריאה ל-`port-discovery.discoverActiveTargets()` שמחזירה `[{ port, app, target }, ...]`.
- ה-`tick` רץ port discovery כל ~2s. לכל target חי > attach + inject.
- נשמרת רשימה של targetIds פעילים כדי לא לחזור פעמיים על אותו target.
- לוגים: `[Word] target attached on port 51244`, `[Excel] target detached`.

**שינויים ב-[scripts/tray-icon.ps1](../scripts/tray-icon.ps1) (חלקי - הליבה ב-M3):**
- עדכון תפריט Auto-Enable: כותב `--remote-debugging-port=0` במקום `--remote-debugging-port=9222`.
- migration: בעלייה, אם הערך הקיים הוא `--remote-debugging-port=9222` > מעדכן ל-`0` בלי שאלה (זה הערך שלנו, אנחנו מעדכנים אותו).

**בדיקות:**
- script שמפעיל 3 אפליקציות (Auto-Enable פעיל), ובודק שה-injector זיהה את שלושתן ב-status file.
- בדיקה ידנית: סגירה של Excel באמצע > status מתעדכן, Word ו-PowerPoint נשארים מחוברים.

### M2 - Tray מצטבר (אגרגציה ב-N אפליקציות)
שדרוג של tray-icon.ps1 לדווח על מספר משתנה של אפליקציות פתוחות.

**שינויים ב-[scripts/tray-icon.ps1](../scripts/tray-icon.ps1):**
- מבנה status file חדש (JSON):
  ```json
  {
    "discovered": [
      {"app": "Word", "port": 51244, "status": "CONNECTED"},
      {"app": "Excel", "port": 51891, "status": "CONNECTED"}
    ],
    "lastDiscovery": "2026-04-25T10:00:00Z",
    "errors": []
  }
  ```
  אם לא ניתן לפרסר > fallback לקריאה כמחרוזת ישנה (תאימות אחורה ל-v0.1.x).
- תפריט ימני:
  - Status > מציג שורה לכל אפליקציה שהתגלתה (`Word: CONNECTED`, `Excel: CONNECTED`).
  - Auto-Enable - **נשאר**, רק עם ערך מעודכן.
  - Re-discover (force tick) - אופציה ידנית להפעיל port discovery מיד.
  - Launcher... (פותח את ה-launcher האופציונלי, אם הותקן).
  - Uninstall.
- צבע האייקון:
  - ירוק: לפחות Office אחד פתוח עם Claude מחובר.
  - צהוב: Office פתוח אבל ללא Claude target (משתמש לא פתח את הפאנל).
  - אדום: שגיאת port discovery (tasklist/netstat נכשלים).
  - אפור: אין Office פתוח.
- האייקון: "C" קטן (Claude), שומר על האסתטיקה הקיימת.

### M3 - Launcher UI (אופציונלי, nice to have)
לאחר שהליבה עובדת, ה-launcher הוא תוספת UX לאלה שמעדיפים כניסה מסורתית.

**קבצים חדשים:**
- `scripts/launcher.ps1` - חלון WinForms עם 4 כפתורים (Word / Excel / PowerPoint / All). לוגואי Microsoft 365 רשמיים. לוגו Claude. RTL.
- `launcher.vbs` - wrapper שקט.
- `scripts/create-launcher-shortcut.ps1` - יוצר קיצור על Desktop וב-Start Menu.

**הערה:** ה-launcher לא נדרש כדי שהמערכת תעבוד. הוא רק חוסך מהמשתמש לפתוח Office בעצמו - הוא פותח אותו עם בחירת קובץ אם רוצה.

### M4 - File associations (אופציונלי, OFF as default)
**הוחלט ב-2026-04-25:** ב-install תופיע שאלת Y/N עם ברירת מחדל = N. אפשר להפעיל גם מאוחר יותר דרך תפריט ה-tray.

**קבצים חדשים:**
- `scripts/associations.ps1 -Install|-Uninstall` - רישום ב-HKCU\Software\Classes\Applications כאפשרות "Open with" (לא ברירת מחדל).

### M5 - תיעוד ו-screenshots
**עדכונים:**
- `README.md` + `README.he.md` - "Works with Word, Excel, PowerPoint via Auto-Enable. One install, all three apps work everywhere."
- `CHANGELOG.md` - entry ל-v0.2.0. אין breaking changes (Auto-Enable עדכן את הערך, לא הוסר).
- `docs/screenshots-plan.md` - 3 screenshots חדשים (Excel/PowerPoint לפני/אחרי, status בtray עם 3 אפליקציות).
- `docs/security.md` - עדכון: "WebView2 debug listens on dynamic ports (49152-65535) per Office process. Localhost only."
- `scripts/_run-tests-internal.js` - מוסיף בדיקות mock של 3 PIDs ב-tasklist + 3 ports ב-netstat.

### M6 - Migration ו-uninstall נקי
**שינויים ב-[install.bat](../install.bat):**
- בעלייה מ-v0.1.x: זיהוי `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222` > עדכון ל-`=0` אוטומטית (זה הערך שלנו).
- רישום `ClaudeOfficeRTL` במקום `ClaudeWordRTL` ב-HKCU\Uninstall.

**שינויים ב-[uninstall.bat](../uninstall.bat):**
- מחיקת `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` רק אם ערכו הוא `--remote-debugging-port=0` (הערך שלנו) או `--remote-debugging-port=9222` (הערך הישן שלנו). שמירה על ערכים שהמשתמש ערך.

**שינויים ב-[scripts/check-update.js](../scripts/check-update.js), [check-update.bat](../check-update.bat):**
- אם זוהתה גרסה 0.1.x > הצגת הודעת migration אוטומטי בעלייה, ללא שבירה.

---

## 6. שינויים לפי קובץ (רשימה מרוכזת)

| קובץ | שינוי | M |
|-------|-------|---|
| `probe/launch-office-dynamic.bat` | **חדש** (קיים) | M0 |
| `probe/dynamic-port-discovery.js` | **חדש** (קיים) | M0 |
| `lib/office-apps.js` | **חדש** - metadata רזה (name, processName, urlHostInfo). ללא ports | M1 |
| `scripts/port-discovery.js` | **חדש** - מודול tasklist+netstat+probe (production-grade) | M1 |
| `scripts/inject.js` | החלפת PORT קבוע ב-port-discovery דינמי, attach לכל target חי | M1 |
| `scripts/tray-icon.ps1` | status JSON חדש (discovered[]), 3 צבעים, Auto-Enable עם ערך מעודכן | M2 |
| `scripts/launcher.ps1` | **חדש (אופציונלי)** - UI בחירת אפליקציה עם לוגואי MS 365 | M3 |
| `launcher.vbs` | **חדש (אופציונלי)** - wrapper שקט | M3 |
| `scripts/create-launcher-shortcut.ps1` | **חדש (אופציונלי)** | M3 |
| `scripts/associations.ps1` | **חדש (אופציונלי, OFF default)** | M4 |
| `install.bat` | בדיקת זמינות 3 EXEs, Auto-Enable עם `port=0`, רישום `ClaudeOfficeRTL`, שאלת associations | M4+M6 |
| `README.md` | עדכון מקיף - "Works with Word/Excel/PowerPoint via single Auto-Enable" | M5 |
| `README.he.md` | עדכון מקיף | M5 |
| `CHANGELOG.md` | entry v0.2.0, ללא breaking changes | M5 |
| `docs/screenshots-plan.md` | screenshots חדשים | M5 |
| `docs/security.md` | dynamic ports 49152-65535, localhost only | M5 |
| `scripts/_run-tests-internal.js` | בדיקות mock של tasklist+netstat | M5 |
| `uninstall.bat` | מחיקת env var רק אם ערכו שלנו (`port=0` או `port=9222`) | M6 |
| `cleanup.bat` | שחרור lock/pid (יחיד) | M6 |
| `doctor.bat` | דיאגנוסטיקה: tasklist, netstat, port discovery dry-run | M6 |
| `start.bat` | מפעיל launcher אם הותקן, אחרת הוראות הפעלה ידנית | M6 |
| `check-update.bat`, `scripts/check-update.js` | migration hint מ-v0.1.x | M6 |
| `package.json` (root + scripts/) | version bump, תיאור מעודכן | M6 |

**סה"כ:** 9 קבצים חדשים (מתוכם 5 אופציונליים), 11 קבצים קיימים משודרגים. **אין יותר wrappers נפרדים לכל אפליקציה.**

---

## 7. סיכוני עבודה

| סיכון | השפעה | מיגור |
|-------|--------|-------|
| `tasklist`/`netstat` נכשלים על מערכת מסוימת (גרסת Windows ישנה, AV חוסם) | port discovery לא עובד | M0 POC חייב לאמת על המכונה של אסף; doctor.bat יבדוק ידנית |
| `tasklist /FO CSV` בעברית/שפה אחרת מחזיר טקסט מתורגם | regex parsing נכשל | להשתמש ב-`/NH` (כבר נעשה), regex על מספרים בלבד; בדיקה על Windows עברית |
| מספר WebView2 PIDs לכל Office app (browser, GPU, renderer, utility) | רעש בלוג, "Found N PIDs" מטעה | להבהיר בלוג שרק תהליך browser מאזין על TCP; השאר נסננים בשקט |
| Claude add-in ב-Excel/PowerPoint אולי משתמש ב-CSS selectors שונים ב-edge cases (pivot tables, slides) | RTL שבור בחלקים | בדיקה ויזואלית ב-M1, הוספת selectors ספציפיים אם צריך |
| WebView2 debug ports פתוחים בו-זמנית בטווח dynamic | תיאורטית surface ל-attacks מקומיים | localhost-only (כבר היה), עדכון docs/security.md |
| Auto-Enable ישן (`port=9222`) נשאר אצל משתמשים בעלייה | רק Word עובד, השאר נכשלים | install.bat מזהה ומעדכן ל-`port=0` אוטומטית (זה הערך שלנו) |
| Race: Office נפתח אבל ה-tick הבא של ה-injector עוד 2s | RTL מתעכב 2-3 שניות מהפתיחה | מקובל; אפשר לקצר POLL_MS ל-1s אם מטריד |
| AV מסמן `tasklist`+`netstat` ב-injector כפעולה חשודה | התקנה נכשלת אצל לקוחות | ניסוח docs/security.md מפורש, חתימה דיגיטלית בעתיד |

---

## 8. תוכנית בדיקות

### ידני (לפני release)
1. **M0:** הרצת `probe/launch-office-dynamic.bat` ואז `node probe/dynamic-port-discovery.js`. תוצאה: `SUCCESS` עם 3 אפליקציות מזוהות.
2. התקנה נקייה על מכונה בלי v0.1.x קודם > Auto-Enable פעיל אוטומטית עם `port=0`.
3. שדרוג ממכונה עם v0.1.x מותקן + auto-enable=9222 מופעל > מעבר חלק ל-`port=0`.
4. **התרחיש המרכזי:** Word + Excel + PowerPoint נפתחים דרך taskbar רגיל (לא דרך launcher). RTL עובד בכל שלוש.
5. פתיחת Word בלבד > tray ירוק עם "1 of 1 connected". פתיחת Excel נוסף > "2 of 2". סגירת Word > "1 of 1" (Excel בלבד).
6. סגירת כל האפליקציות > tray אפור.
7. file associations ON (אופציונלי): double-click על .xlsx > Excel נפתח, RTL עובד.
8. Uninstall מלא > בדיקת רישום, env var (רק אם הערך שלנו), קיצורים.

### אוטומטי
- `scripts/_run-tests-internal.js` - mock של פלט `tasklist` ו-`netstat` (טקסט קבוע) > בדיקה ש-port-discovery מחלץ נכון.
- mock CDP server שמחזיר `_host_Info=Word/Excel/Powerpoint` > בדיקת זיהוי app.
- CI (אם נוסיף): GitHub Actions עם Node על Windows, הרצת port-discovery מול mocks.

---

## 9. תוכנית Release

- **0.1.x** - ממשיך לקבל תיקוני אבטחה בלבד
- **0.2.0-alpha** - פריסה פנימית אצל אסף לבדיקה שבוע
- **0.2.0-beta** - פריסה לעד 5 משתמשי early adopter עם feedback
- **0.2.0** - release ציבורי, blog post / LinkedIn post על "Claude for Office RTL"

---

## 10. החלטות

**סגורות (2026-04-25):**
- **Outlook:** לא נתמך. לא מתייחסים אליו בפרויקט הזה.
- **Launcher UI - לוגואים:** משתמשים בלוגואי Microsoft 365 (Word/Excel/PowerPoint הרשמיים). יש לוודא שימוש מותר לפי Microsoft Brand Guidelines - לא לערוך, לא לעוות, רק להציג בהקשר של אינטגרציה לגיטימית.
- **ארכיטקטורה: dynamic ports.** `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0`. כל אפליקציית Office תופסת פורט פנוי משלה ב-49152-65535. ה-injector מגלה דרך `tasklist`+`netstat`, מזהה אפליקציה לפי `_host_Info=`. אין יותר פורטים קבועים (9222/9223/9224) ולא wrappers נפרדים.
- **Auto-Enable נשאר.** רק עם ערך מעודכן (`port=0` במקום `port=9222`). זה הופך לנתיב הראשי - לחיצה אחת, כל פתיחה של Office עובדת.
- **File associations:** OFF as default. בהתקנה תופיע שאלת Y/N עם ברירת מחדל = N. נחשוף "Enable file associations" גם בתפריט ה-tray כך שאפשר להפעיל בכל רגע אחרי ההתקנה.
- **שיטת עבודה:** קודם כל בדיקה מקומית מקצה לקצה. רק אחרי שהכול עובד - מעלים ל-GitHub כ-release ציבורי.

**סגורות נוספות (2026-04-25):**
- **Tray menu - אופציה B (מינימליסטי).** התפריט הימני יכלול: Auto-Enable toggle, File associations toggle, Show injector log, Check for updates, Uninstall, Exit tray. **אין רשימת אפליקציות מחוברות בתפריט עצמו** - הסטטוס משתקף בצבע האייקון (ירוק/צהוב/אדום/אפור) וב-tooltip במעבר עכבר ("X of Y Office apps connected"). הולם את הפילוסופיה של הפרויקט: פשוט, שקט, "פשוט עובד".

**Auto-Enable Migration - אופציה C (silent + info notice):** install.bat של v0.2.0 בודק אם הערך הקיים הוא `--remote-debugging-port=9222`. אם כן > מעדכן ל-`port=0` ללא שאלה (זה הערך שלנו), אבל מציג הודעת info בקונסול ובלוג: `[INFO] Auto-Enable updated to dynamic ports (was fixed port 9222) for Excel + PowerPoint support`. אם הערך שונה או לא קיים > לא נוגע בו.

**שם המוצר - "Claude for Office RTL Fix":** שינוי מינימלי מ"for Word" ל"for Office". DisplayName ב-Apps and Features, README, CHANGELOG, ב-tray title - הכל מתעדכן. **GitHub repo URL נשאר `asaf-aizone/Claude-for-word-RTL-fix`** (לא משנים URL - שובר external links, מנגנון check-update.js, וקיצורים שמשתמשים שמרו). אם בעתיד תהיה הצדקה גדולה - אפשר לשקול rename של ה-repo עם redirect אוטומטי של GitHub. כרגע לא נחוץ.

**כל ההחלטות נסגרו. ניתן להתחיל M1.**

---

## 11. צעד ראשון כשמתחילים

1. **M0 - POC validation.** סגור את כל אפליקציות Office, הרץ `probe/launch-office-dynamic.bat`, פתח Claude add-in בכל אפליקציה, הרץ `node probe/dynamic-port-discovery.js`. דרישה: פלט `SUCCESS`. אם נכשל > תיעוד ב-`docs/bugs/` וחזרה לחלופה D.
2. סיכום 3 ההחלטות הפתוחות הנותרות > עדכון המסמך הזה.
3. פתיחת branch `v0.2.0-office-expansion`.
4. **M1** = הליבה: port-discovery production-grade ב-`scripts/`, שילוב ב-`inject.js`, עדכון Auto-Enable ל-`port=0`.
5. M2-M6 מקבילים. M3 (launcher) ו-M4 (file associations) אופציונליים, אפשר לדחות אם הזמן לוחץ.
