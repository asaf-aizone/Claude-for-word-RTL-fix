# Bug #002 - install.bat נכשל בשקט (LF line endings)

**תאריך גילוי:** 2026-04-21 10:00
**גרסה:** v0.1.0 (טרום-release)
**מדווח:** asafabram
**חומרה:** קריטית (חוסם release - לא ניתן להתקין את המוצר)
**סטטוס:** סגור ב-v0.1.0 (נבדק 2026-04-22)

## תיקון
שני קבצים נוספו לשורש הריפו:
- `.gitattributes` - כופה CRLF על `.bat`, `.cmd`, `.ps1`, `.vbs`; LF על `.sh`
- `.editorconfig` - כלל `end_of_line = crlf` לקבצי [*.{bat,cmd,ps1,vbs}]

בדיקה: כל 7 קבצי ה-.bat בציבורי כרגע ב-CRLF (עברה אימות ב-`od -c`).

## תיאור קצר

`install.bat` מתחיל לרוץ, כותב שורה אחת לקובץ הלוג, ואז יוצא בשקט
בלי שום output גלוי ובלי להתקין כלום. המשתמש רואה cmd prompt חוזר
מיד בלי שום הודעה. **אף שלב מההתקנה לא מתבצע בפועל** - אין npm install,
אין Startup entry, אין רישום Apps and Features, אין הפעלת tray.

## ניתוח שורש: line endings

כל 7 קבצי ה-`.bat` בפרויקט נשמרו עם **LF only (Unix)** במקום **CRLF (Windows)**:

```
install.bat       : LF ONLY
uninstall.bat     : LF ONLY
start.bat         : LF ONLY
cleanup.bat       : LF ONLY
doctor.bat        : LF ONLY
word-wrapper.bat  : LF ONLY
check-update.bat  : LF ONLY
```

**cmd.exe מתנהג לא consistently עם LF-only:**
- `uninstall.bat` ו-`cleanup.bat` עובדים (בדוק ואומת במהלך הסשן).
- `install.bat` נכשל בשקט.

**הסיבה להבדל:**

| קובץ | שימוש ב-`call :log` subroutine | תוצאה |
|------|-------------------------------|-------|
| install.bat | 51 שימושים | נכשל בשקט |
| uninstall.bat | 0 | עובד |
| cleanup.bat | 0 | עובד |

cmd.exe עם LF endings מצליח לפרסר פקודות פשוטות (`echo`, `set`, `if`,
`pause`, `reg add`) אבל נכשל שקטוט על הקומבינציה של **`setlocal
EnableDelayedExpansion` + `call :label`**. ברגע שמגיע ה-`call :log`
הראשון, cmd לא מוצא את ה-label בצורה תקינה ויוצא מה-script.

## שלבי שחזור

1. לקלון את ה-repo או לחלץ מ-zip.
2. `cd` לתיקייה.
3. להריץ `install.bat` (דאבל-קליק או מחלון cmd).
4. **צפוי:** התקנה מלאה עם output.
5. **בפועל:** חלון cmd נפתח, רץ חצי-שנייה, חוזר לפרומפט בלי שום הודעה.

## ראיות

### קובץ `install.log` אחרי ריצה כושלת

```
Install started Tue 04/21/2026 10:00:04.28
```

(רק שורה אחת. השורה הזו מיוצרת בשורה 21 של `install.bat` באמצעות
redirect רגיל `echo ... > "%LOG%"`. כל השורות הבאות משתמשות ב-
`call :log "..."` ולא מופיעות.)

### מה אמור להופיע ב-install.log אחרי התקנה מוצלחת

לפי הקוד, היו צריכות להופיע כ-30 שורות שכוללות:

```
================================================================
 Claude for Word RTL - Installer
================================================================

[1/4] Checking prerequisites...
  [OK] Node.js found.
  [OK] Word is not running.
  [OK] Word found.

[2/4] Installing npm dependencies...
  Installing dependencies, ~15 seconds...
  [OK] Dependencies installed.

[3/4] Creating Startup folder entry...
  [OK] Startup entry created.

[4/4] Registering with Apps and Features...
  [OK] Registered.

  Starting tray icon...
  [OK] Tray icon started.

================================================================
 Installation complete.
================================================================
...
```

### בדיקת מצב המערכת אחרי "התקנה" כושלת

```
node_modules?            False  (לא הותקן)
Startup entry?           False  (לא נוצר)
Tray running?            False  (לא הופעל)
Apps and Features entry? False  (לא נרשם)
```

## פתרון - תיקון הקבצים

### הפתרון המיידי

להמיר את כל קבצי `.bat` ל-CRLF. פקודה אחת ב-PowerShell:

```powershell
Get-ChildItem "C:\Users\asafa\Downloads\claude-word-rtl\*.bat" | ForEach-Object {
    $content = Get-Content $_.FullName -Raw
    $normalized = ($content -replace "`r`n", "`n") -replace "`n", "`r`n"
    [IO.File]::WriteAllText($_.FullName, $normalized)
}
```

כדאי גם להריץ על `scripts\*.vbs` ו-`scripts\*.ps1` אם הם עם אותה בעיה.

### מניעה קבועה - `.gitattributes`

להוסיף קובץ `.gitattributes` לשורש ה-repo עם:

```
*.bat text eol=crlf
*.cmd text eol=crlf
*.ps1 text eol=crlf
*.vbs text eol=crlf
```

זה מבטיח ש-Git יאחסן ויוציא את הקבצים עם CRLF ללא קשר ל-autocrlf
של המשתמש או הסביבה בה נערכו.

### מניעה נוספת - editor config

להוסיף `.editorconfig`:

```ini
root = true

[*.bat]
end_of_line = crlf

[*.cmd]
end_of_line = crlf

[*.ps1]
end_of_line = crlf

[*.vbs]
end_of_line = crlf
```

זה מבטיח שעורכי טקסט (VSCode, Sublime, Notepad++) ישמרו עם CRLF.

## בדיקה אחרי התיקון

1. להמיר את כל קבצי ה-.bat ל-CRLF (ראו פקודה למעלה).
2. להריץ `install.bat` > אמור להופיע 4 שלבים עם output גלוי, npm install
   של ~15 שניות, Startup entry, tray מופעל.
3. להריץ `uninstall.bat` > אמור לעבוד כרגיל (כבר עובד היום).
4. להריץ `start.bat` > קיים חשד שגם הוא סובל מהבעיה (לא נבדק במישרין
   בסשן זה, אבל יש לו לוגיקה מורכבת דומה). לוודא.

## השפעה על התקנה ציבורית

**זהו blocker מלא לפרסום v0.1.0.** כל משתמש שיוריד את הזיפ ויריץ
install.bat יחווה בדיוק את מה שאנחנו חווינו: חלון cmd נפתח, נסגר תוך
חצי שנייה, כלום לא קורה. זה יתן חוויה של "המוצר לא עובד" או "המוצר
שבור" ויפגוש קהל היעד בתחושה הפוכה ממה שאנחנו מנסים להשיג.

**חובה לתקן ולבדוק התקנה על מחשב "נקי" (VM/חבר/machine אחר) לפני
release.**

## הקשר לבאגים אחרים

- **Bug #001** (Disconnect greyed-out) - לא קשור ישירות, אבל שני
  הבאגים האלו מחזקים המלצה חזקה: לפני v0.1.0 להוציא **עמוד QA עצמאי**
  שבודק את התרחישים הבסיסיים על VM נקי: התקנה, Connect, Disconnect,
  Auto-enable on/off, uninstall. כל אחד מאלו נכשל היום על המכונה של
  המפתח או על מכונה נקייה.
- **שבירת הסשן הנוכחי לצילומי מסך:** הבאג הזה מנע את יצירת צילום
  ה-`installer-done.png` שהוא אחד מצילומי ה-README. עד שהבאג ייסגר,
  לא ניתן לצלם התקנה מוצלחת. ניתן (כפתרון ביניים) לייצר צילום מסך
  סינתטי מהטקסט שהיה אמור להופיע, אבל זה לא פתרון ראוי ל-README
  של release.

## Checklist לסשן הקוד

- [ ] להריץ את סקריפט ההמרה (PowerShell למעלה) על כל קבצי .bat ו-.cmd.
- [ ] לבדוק שגם scripts/*.vbs ו-scripts/*.ps1 נמצאים ב-CRLF.
- [ ] לבדוק שגם קבצי config אחרים (.json, .md) לא נשברו.
- [ ] להוסיף `.gitattributes` לשורש.
- [ ] להוסיף `.editorconfig` לשורש.
- [ ] לעשות commit של התיקונים: "Fix line endings on .bat scripts (CRLF)
      - fixes silent failure of install.bat on Windows".
- [ ] לבדוק התקנה מקצה לקצה על machine נוסף (לא המכונה של המפתח).
- [ ] לעדכן את סקיל צילומי המסך - להשלים את `installer-done.png`.
