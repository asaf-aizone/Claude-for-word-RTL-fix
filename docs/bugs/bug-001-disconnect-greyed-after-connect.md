# Bug #001 - Disconnect מופיע כאפור אחרי Connect שכשל

**תאריך גילוי:** 2026-04-21 07:24
**גרסה:** v0.1.0 (טרום-release)
**מדווח:** asafabram
**חומרה:** בינונית (UX מטעה, אין נזק לדאטה)
**סטטוס:** סגור ב-v0.1.0 (נבדק 2026-04-22)

## תיקון
`scripts/tray-icon.ps1:684` - Disconnect מופעל עכשיו כאשר אחד מהשלושה נכון:
```powershell
$miDisconnect.Enabled = ($wordRunning -or $injectorAlive -or $connectInProgress)
```
במקום התנאי הקודם שהיה תלוי רק ב-`$wordRunning`. זה סוגר בדיוק את התרחיש של הבאג - injector חי בלי Word.

## תיאור קצר

אחרי לחיצה על Connect מהתפריט ולחיצת אישור על דיאלוג ההפעלה מחדש,
המשתמש נשאר במצב חסום: התפריט מציג Disconnect באפור (לא ניתן ללחיצה),
למרות ש-Connect כביכול רץ. לא ברור למשתמש איך לצאת מהמצב.

## שלבי שחזור

1. הכלי מותקן, האייקון פעיל. Word סגור.
2. קליק ימני על האייקון > **Connect**.
3. דיאלוג האישור קופץ. לחיצה על **אישור**.
4. (סביר שמשהו נכשל ברקע - Word לא נפתח בפועל, או נפתח ונסגר מיידית.)
5. קליק ימני שוב על האייקון > התפריט מראה את Disconnect באפור.
6. אין אפשרות ללחוץ Disconnect כדי "לאפס" את המצב.

## התנהגות צפויה

אחת משתי האופציות:
- **אם Connect הצליח חלקית** (תהליך injector רץ אבל לא מצא Word): Disconnect צריך להיות פעיל כדי לעצור את ה-injector בלי לעצור Word. לחילופין, האייקון צריך להישאר אדום ועם הצעת recovery.
- **אם Connect נכשל לחלוטין**: המערכת צריכה לחזור למצב ההתחלתי - האייקון אדום, Disconnect מאופר, Connect פעיל. כך שהמשתמש יכול פשוט לנסות שוב.

## התנהגות בפועל

- האייקון: צריך לבדוק (לא תועד בצילום)
- Disconnect: אפור, לא ניתן ללחיצה
- Connect: פעיל
- אפליקציה Word: לא רצה
- ה-injector: רץ (PID 25644 בעת הבדיקה, התחיל ב-07:23:28)
- קובץ status: `DISCONNECTED`
- לוג ה-injector: 5 שורות בלבד, האחרונה: `listTargets failed (tick 1): AggregateError`

## ראיות שנאספו

### צילום מסך
`docs/bugs/bug-001-disconnect-greyed.png`

### תוכן `%TEMP%\claude-word-rtl.log` בעת התקלה

```
2026-04-21T04:23:28.671Z Claude RTL injector starting. Watching localhost:9222
2026-04-21T04:23:28.678Z Match pattern (primary): {}
2026-04-21T04:23:28.680Z Match pattern (fallback): {}
2026-04-21T04:23:28.682Z Ctrl+C to stop.
2026-04-21T04:23:28.704Z listTargets failed (tick 1): AggregateError
```
(הזמן ב-UTC; מקומי = 07:23:28)

ה-`AggregateError` מצביע על כך ש-WebSocket לא הצליח להתחבר ל-`localhost:9222`,
דהיינו אף תהליך WebView2 לא חשף debug port ברגע ה-Connect.

### קבצי סטטוס ב-`%TEMP%`

```
claude-word-rtl.lock      28 bytes  mtime=07:23:28
claude-word-rtl.log      296 bytes  mtime=07:23:28
claude-word-rtl.pid        6 bytes  mtime=07:23:28
claude-word-rtl.status    13 bytes  mtime=07:23:28  (תוכן: "DISCONNECTED")
claude-word-rtl.tray.pid   7 bytes  mtime=06:25:42
```

### מצב WINWORD

אין תהליך `WINWORD.EXE` רץ. כלומר Word לא נפתח כתוצאה מ-Connect, או
שנפתח ונסגר מייד לפני שה-injector הספיק להתחבר.

### תהליכי Node לבדיקה

20 תהליכי `node.exe` בזיכרון, רבים מהם מ-05:43-05:44 (בדיקות מוקדמות).
אפשרי שמדובר ב-leak שלא מקושר ישירות לבאג הנוכחי, אבל ייתכן שריבוי
תהליכי Node ישנים מעורב בבעיה (לדוגמה: lock לא משוחרר). שווה לבדוק.

### משתנה סביבה

```
USER-LEVEL WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: [ריק]
PROCESS  WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: [ריק]
```

**ממצא צדדי:** המשתמש דיווח שלחץ "אישור" על דיאלוג Auto-enable מוקדם
יותר באותה סשן (07:17), אבל **המשתנה לא נמצא ב-HKCU\Environment**.
ייתכן שזה באג נפרד - Auto-enable כביכול מתבצע אבל לא נכתב.
ראו "באגים קשורים" למטה.

## ניתוח שורש אפשרי

הבאג עשוי לנבוע מאחד מהבאים, או משילוב:

1. **State machine לא מטפל בכישלון של Word relaunch**: אם ה-wrapper
   רץ אבל Word לא נפתח (אנטי-וירוס חוסם, EDR, או סתם Word זומבי
   אחר), ה-state machine נכנס למצב intermediate שבו ה-injector רץ
   (ולכן status נכתב כ"DISCONNECTED" במקום היחזרה ל-IDLE) אבל
   Disconnect מותנה ב-`isWordRunning()` ולכן מאופר.
2. **בדיקת זמינות של Disconnect מבוססת על process check ולא על
   status file**: ייתכן שה-tray בודק קיום של תהליך WINWORD.EXE כדי
   להפעיל את Disconnect. כאן אין WINWORD ולכן Disconnect מאופר -
   למרות שהמצב הנכון הוא "Connect חצי-תקוע, יש injector חי".
3. **Lock לא משוחרר**: קובץ ה-lock ב-`%TEMP%` נמצא, וה-injector
   חושב שהוא בעלים. הניסיון שוב להריץ Connect (אם המשתמש ינסה)
   עלול להיכשל בשקט.

## הצעות לתיקון

(לסשן הקוד - לא לסשן הזה)

1. **Disconnect צריך להיות פעיל כש-injector רץ**, גם אם Word לא נמצא.
   הפעולה תעצור את ה-injector ותחזיר את ה-tray למצב IDLE.
2. **תוסיף timeout ל-Connect**: אם תוך 30 שניות ה-injector לא מצליח
   להתחבר, להחזיר את המצב ל-IDLE אוטומטית עם notification למשתמש.
3. **לוג מפורט יותר**: לפחות לרשום ניסיונות `listTargets` חוזרים, לא
   רק את הראשון, כדי שהמשתמש יראה אם הבעיה רגעית או תמידית.
4. **כפתור recovery בתפריט**: "Reset state" שמאפס את כל קבצי ה-state
   ב-`%TEMP%` בלי uninstall.

## פתרון עוקף לזמן הזה

1. סגור Word ידנית אם פתוח.
2. הרץ `cleanup.bat` מתיקיית הפרויקט - יעצור את ה-injector ויאפס.
3. (אופציונלי) Exit מה-tray + Restart של ה-tray (`start.bat`).
4. נסה שוב Connect מחדש.

## באגים קשורים (לחקור בנפרד)

- **Auto-enable - UX feedback חסר (עודכן 2026-04-21 אחרי הפסקה)**:
  בבדיקה חוזרת לאחר ~30 דקות, המשתנה `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`
  *כן* נמצא ב-`HKCU\Environment` עם ערך `--remote-debugging-port=9222`.
  כלומר Auto-enable כן כתב בפועל. הבעיה המקורית הייתה:
  (א) בבדיקה מיידית אחרי הלחיצה המשתנה היה ריק - מה שמרמז על delay
  או async write שלא מסונכרן עם ה-UI feedback, או
  (ב) לא היה סימון visual ב-checkbox בתפריט אחרי ההפעלה, מה שגרם
  למשתמש לחשוב שהפעולה נכשלה.
  **המלצה:** להוסיף re-render מפורש של ה-checkmark בתפריט מיד אחרי
  כתיבת המשתנה, וגם לוודא שהכתיבה synchronous (או לפחות לא לסגור את
  ה-handler לפני שה-registry write מסתיים).
- **20 תהליכי Node תקועים**: לבדוק האם הם זומבים מסשנים קודמים, ואם
  כן - האם `cleanup.bat` באמת תופס את כולם או רק את אלו עם PID
  פעיל ב-`%TEMP%\claude-word-rtl.pid`.

## הקשר לטיוטת ה-README

הטיוטה כרגע (`docs/README-landing-draft.md`, סקציית Troubleshooting)
מכסה מקרה של "Icon stays red after Connect" - אבל זה תרחיש שונה
(האייקון אדום והמשתמש יודע מה לעשות). הבאג הזה משאיר את המשתמש
במצב לא-מוגדר. אחרי שהבאג ייסגר, כדאי לוודא שיש שורת troubleshooting
לתרחיש הזה.
