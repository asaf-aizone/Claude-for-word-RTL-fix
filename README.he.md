<div dir="rtl">

# Claude for Office RTL Fix (Word, Excel, PowerPoint, Outlook)

[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue)](LICENSE)
[![Node](https://img.shields.io/badge/node-%E2%89%A516-brightgreen)](https://nodejs.org/)
[![No Telemetry](https://img.shields.io/badge/telemetry-none-success)](#פרטיות)
[![Local Only](https://img.shields.io/badge/network-localhost%20only-success)](#הערת-אבטחה)

**גרסה מלאה (דו-לשונית, עם צילומי מסך ותמונות): [README.md](README.md)** · **יומן שינויים: [CHANGELOG.md](CHANGELOG.md)**

תיקון CSS וטיפוגרפיה בצד הלקוח לתצוגה העברית בתוסף Claude ל-Microsoft Word, Excel, PowerPoint ו-Outlook. החל מ-v0.2.0 הכלי תומך ב-Word/Excel/PowerPoint בו-זמנית. **v0.3.0 הוסיף את Outlook הקלאסי** מאחורי מודל opt-in מוקשח (אישור מפורש לכל הפעלה, ניתוק אוטומטי אחרי 15 דקות, redaction של ה-URL בלוג). שם המאגר ב-GitHub עודכן ב-v0.2.1 ל-`Claude-for-Office-RTL-fix` (היה `Claude-for-word-RTL-fix`). GitHub מחזיק redirect קבוע מהשם הישן, אז clones ו-bookmarks מגרסאות קודמות ממשיכים לעבוד.

תוסף Claude הרשמי ל-Office מציג כיום טקסט עברי משמאל לימין, עם סימני רשימה ופיסוק בצד הלא נכון. הכלי הזה מתחבר לחלונית WebView2 של התוסף בכל אחת מארבע האפליקציות באמצעות Chrome DevTools Protocol הסטנדרטי, ומזריק גיליון סגנונות וכן MutationObserver קטן כדי לתקן את התצוגה.

הכלי קיים מטעמי נגישות: דוברי עברית זקוקים לתצוגת RTL כדי לקרוא את תשובות Claude בחלונית. בלי התיקון העברית מוצגת LTR עם טיפול bidi שבור, והפאנל בפועל לא שמיש. מדובר בהתאמת נגישות ליציאה שמוצגת מקומית, בדומה לגיליון סגנונות של משתמש (Stylus/Stylish), קורא מסך או מזריק dark-mode. השירות עצמו לא משתנה בשום צורה.

הכל פועל מקומית במחשב שלך. שום דבר לא נשלח ברשת.

> **Windows בלבד.** הכלי לא עובד על macOS או Linux. תוסף Claude ל-Office מבוסס על WebView2 של מיקרוסופט, שקיים רק ב-Windows. ל-Office ל-Mac יש runtime אחר (WKWebView) שלא חושף את אותו debugging interface, וכל שכבת ההפעלה (bat, vbs, PowerShell, Registry, Startup folder) לא רלוונטית שם. אם אתם על Mac, אין port מ-Office.

> **אזהרה למחשבים מנוהלי-ארגון.** הכלי מתחבר ל-Microsoft Word דרך Chrome DevTools Protocol ומזריק JavaScript לתוך WebView2, ומפעיל את עצמו דרך VBS hidden launcher ו-PowerShell. הצירוף הזה דומה מבחינה מבנית לטכניקות שגונבי-מידע משתמשים בהן, ולכן מערכות EDR ארגוניות (Microsoft Defender for Endpoint, CrowdStrike Falcon, SentinelOne, Sophos) עלולות לזהות את ההתקנה כפעילות חשודה ולנתק את המכונה מהרשת (host isolation) באופן אוטומטי. **אין להתקין על מחשב מנוהל-ארגון בלי אישור מקדים מצוות אבטחת המידע** ובלי הוספת ה-hash וה-path של הקבצים ל-allowlist. המחבר אינו אחראי לתגובות מערכות אבטחה ארגוניות.

> **אזהרה למחשבים מנוהלי-ארגון.** הכלי מתחבר ל-Microsoft Word דרך Chrome DevTools Protocol ומזריק JavaScript לתוך WebView2, ומפעיל את עצמו דרך VBS hidden launcher ו-PowerShell. הצירוף הזה דומה מבחינה מבנית לטכניקות שגונבי-מידע משתמשים בהן, ולכן מערכות EDR ארגוניות (Microsoft Defender for Endpoint, CrowdStrike Falcon, SentinelOne, Sophos) עלולות לזהות את ההתקנה כפעילות חשודה ולנתק את המכונה מהרשת (host isolation) באופן אוטומטי. **אין להתקין על מחשב מנוהל-ארגון בלי אישור מקדים מצוות אבטחת המידע** ובלי הוספת ה-hash וה-path של הקבצים ל-allowlist. המחבר אינו אחראי לתגובות מערכות אבטחה ארגוניות.

## מה הכלי עושה

- קובע `direction: rtl` ויישור טקסט מותאם לעברית בחלונית
- מתקן סימני רשימה ממוספרים ובלתי ממוספרים כך שיופיעו מימין
- משאיר בלוקי קוד משמאל לימין לשם נכונות
- מחליף em-dash (—) ו-en-dash (–) במקף קצר (-) בטקסט המוצג
- מחליף חצים (&rarr; &larr; ↔) בפסיקים בטקסט המוצג
- מוחל מחדש אוטומטית אם החלונית נטענת מחדש

## מה הכלי לא עושה

- לא משנה את מה שנשלח ל-Claude
- לא משנה את מה ש-Claude מחזיר מהשרת, רק את אופן התצוגה המקומית
- לא עוקף מגבלות שימוש, מנגנוני הגנה או כל הגבלה אחרת
- לא מבצע הנדסה לאחור של המודל, ה-API או התוסף
- לא אוסף ולא מושך מידע מ-Claude או מתוכן השיחה
- לא מספק גישה אוטומטית ל-Claude. המשתמש מנהל כל שיחה באופן ידני
- לא עוקף את מודל האבטחה של תוספי Microsoft. תוסף Office רץ ללא שינוי על ידי Word, Excel או PowerPoint, בדיוק כפי שאנת'רופיק מספקת אותו
- לא שולח טלמטריה או תעבורת רשת משלו
- לא שומר אישורים, תוכן שיחות או כל מידע אחר
- לא משנה שיוכי קבצים של Office
- לא יוצר משימות מתוזמנות או שירותי רקע
- לא משנה את `Normal.dotm` או כל תבנית אחרת של Word, Excel או PowerPoint
- לא משנה קבצים מחוץ לתיקייה של עצמו, למעט קיצור יחיד בתיקיית Startup לכל משתמש (שמפעיל את אייקון המגש בכניסה למערכת) ומפתח רישום יחיד תחת `HKCU\...\Uninstall\ClaudeWordRTL` (כדי שהכלי יופיע ב-Windows Settings > Apps). `uninstall.bat` מסיר את שניהם

## פרטיות

הכלי הזה פועל לחלוטין על המחשב שלך. ובפירוט:

- **אין טלמטריה, אנליטיקס או מעקב שימוש** מכל סוג שהוא.
- **אין חיבורי רשת יוצאים** מיוזמת הקוד של הכלי.
- **אין איסוף, שמירה או תיעוד** של הפרומפטים שלך, תשובות Claude, טיוטות, תוכן מסמכים או כל נתון אחר מהחלונית. (הכלי אכן כותב קבצי עזר מקומיים קטנים: מזהה התהליך והסטטוס של ה-injector ב-`%TEMP%`, וקבצי לוג אופציונליים של install/doctor לצד הסקריפטים. אף אחד מאלה אינו מכיל תוכן שיחה.)
- **אין שירותי צד שלישי**, הכלי לא פונה לשום שרת של המחבר או של מישהו אחר.
- השיחות שלך עם Claude ממשיכות לעבור ישירות בין WebView2 של Word ל-Anthropic, בדיוק כפי שהיו זורמות בלי הכלי. הכלי רק משנה את אופן התצוגה המקומי של החלונית, הוא לא מתווך, לא משקף, לא בוחן ולא מעביר את הנתונים שלך.

הדברים היחידים שנכתבים לדיסק הם הקבצים של הכלי עצמו בתיקיית ההתקנה, קבצי עזר של סטטוס/PID בתיקיית `%TEMP%`, קיצור יחיד לכל משתמש בתיקיית Startup (`Claude for Word RTL Tray.lnk`) שמפעיל את אייקון המגש בכניסה למערכת, ומפתח יחיד תחת `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL` כדי שהכלי יופיע ב-Windows Settings > Apps > Installed apps. כל אלה מוסרים על ידי `uninstall.bat`. לא מתבצע שינוי נוסף ב-registry.

## הערת אבטחה

בזמן ש-Word, Excel, PowerPoint או Outlook פועלים דרך הכלי הזה, ה-WebView2 של אותה אפליקציה פותח פורט דיבאג על localhost בפורט דינמי (אחד לכל תהליך WebView2 host של Office; ב-v0.2.0 משתמשים ב-`--remote-debugging-port=0`, במקום `9222` הקבוע מ-v0.1.x). משמעות הדבר שכל תהליך מקומי אחר במחשב שלך יכול להתחבר ל-DOM של חלונית Claude (לקרוא טיוטות, עוגיות סשן וכדומה). הפורט הוא localhost בלבד, לא חשוף לרשת, אבל הוא לא דורש אימות.

**Outlook ספציפית (חדש ב-v0.3.0):** כשמשתמש מבקש מ-Claude לסכם מייל או לנסח תגובה, תוכן המייל הופך לחלק מה-DOM של הפאנל. ה-CDP attach חושף אותו לכל תהליך מקומי. החשיפה צרה בזמן (רק כשהפעולה רצה) אבל הקטגוריה רגישה יותר - לכן Outlook מוגן בארבע שכבות שלא פעילות עבור Word/Excel/PowerPoint: (1) opt-in מפורש לכל הפעלה דרך דיאלוג עם Cancel כברירת מחדל ממוקדת, (2) אין auto-launch של ה-injector אם רק Outlook פתוח, (3) ניתוק אוטומטי אחרי 15 דקות, (4) URL redaction בלוג. בנוסף יש פריט תפריט ייעודי **Disconnect Outlook only** שמנתק רק את Outlook בלי לפגוע בשאר. ראו [docs/security.md](docs/security.md#outlook-specific-risks-and-mitigations) לפרטים.

המלצות:

- סגרו את האפליקציה (Word/Excel/PowerPoint) כשאתם לא משתמשים ב-Claude באופן פעיל.
- ב-Outlook: השתמשו ב-Disconnect Outlook only בסוף סשן מייל, אל תסמכו רק על הטיימר של 15 דקות.
- אל תריצו את הכלי על מחשבים משותפים או מחשבים עם תוכנות לא מהימנות.
- במחשבים מנוהלים ארגונית (EDR, DLP), בדקו תחילה עם ה-IT.

ראה [SECURITY.md](SECURITY.md) למודל האיומים המלא ולתהליך דיווח פגיעויות.

## דרישות

- **Windows 10 או 11 בלבד.** macOS ו-Linux לא נתמכים (ראו למעלה).
- Microsoft Office (דסקטופ) - לפחות אחת מבין Word, Excel, PowerPoint, או Outlook הקלאסי, עם התוסף Claude מותקן. תמיכת Outlook היא opt-in; מי שלא רוצה לחבר את Outlook יכול להתעלם מהפיצ'ר לחלוטין. New Outlook (`olk.exe`) לא נתמך
- [Node.js](https://nodejs.org/) 16 או חדש יותר (מותקן ונמצא ב-PATH)

## התקנה

1. שכפל את המאגר או הורד כ-ZIP.
2. סגרו את Microsoft Word/Excel/PowerPoint אם אחת מהן פתוחה (המתקין יבדוק ויתריע).
3. הפעל בלחיצה כפולה על `install.bat`.
   - בהפעלה הראשונה הוא מתקין את `chrome-remote-interface` דרך npm.
   - לוג התקנה מלא נכתב לקובץ `install.log` לצד הסקריפט.
   - Windows SmartScreen עלול להתריע. לחץ על "מידע נוסף" ואז "הפעל בכל זאת" אם אתה סומך על המקור.
   - לא נדרשות הרשאות מנהל. המתקין יוצר קיצור יחיד בתיקיית Startup (למגש) ומפתח רישום יחיד תחת `HKCU\...\Uninstall\ClaudeWordRTL` כדי שהכלי יופיע ב-Windows Settings > Apps.

### איך משתמשים

> **התזרים הנכון (Connect קודם, אז פותחים את הקובץ):**
> 1. אם Word / Excel / PowerPoint / Outlook פתוחים - שמרו וסגרו אותם.
> 2. קליק ימני על אייקון ה-tray ליד השעון, ובוחרים `Connect Word` (או Excel / PowerPoint / Outlook).
> 3. מאשרים בדיאלוג. ה-tray מפעיל את האפליקציה (ריקה) דרך ה-wrapper, והאייקון הופך לירוק.
> 4. **רק עכשיו** פותחים את הקובץ מתוך האפליקציה (File > Open, Recent, או גרירה לחלון). הפאנל יעלה RTL.
>
> **למה הסדר חשוב:** דגל ה-debug של WebView2 נקלט רק בעליית התהליך של Office. קובץ שנפתח *אחרי* שהאפליקציה הופעלה דרך ה-wrapper - יקבל RTL. קובץ שנפתח לפני - לא, וצריך לסגור ולהתחיל שוב.

<details>
<summary>אם האפליקציה כבר פתוחה עם מסמך לא שמור (מסלול fallback)</summary>

ה-tray תומך גם בלחיצת Connect כשהאפליקציה כבר פתוחה - הוא ינסה לסגור אותה בעדינות, יפעיל מחדש דרך ה-wrapper, ויפתח מחדש את אותם מסמכים. המסלול הזה עובר דרך COM enumeration ויש בו יותר חלקים זזים, ולכן הוא **לא מומלץ כברירת מחדל** - הוא קיים בעיקר כדי לא להיתקע אם כבר התחלתם לעבוד. אם משהו נכשל (האפליקציה לא עלתה מחדש, או עלתה בלי המסמכים) - סגרו ידנית והשתמשו בתזרים המומלץ למעלה.

</details>

אייקון המגש הוא נקודת הכניסה היחידה:

1. פתחו את Word, Excel, PowerPoint או Outlook כרגיל (דרך אייקון האפליקציה, קיצור דרך, מסמך, כל מה שאתם רגילים להשתמש בו).
2. לחצו קליק ימני על אייקון המגש ליד השעון ובחרו **Connect Word** / **Connect Excel** / **Connect PowerPoint** / **Connect Outlook** לפי האפליקציה שפתוחה.
3. עבור Word/Excel/PowerPoint: המגש סוגר בצורה מנומסת את האפליקציה, מפעיל אותה מחדש דרך ה-wrapper המתאים עם דגל הדיבאג, ופותח מחדש את המסמכים/חוברות העבודה/מצגות שהיו פתוחים. עברית בפאנל של Claude עכשיו מוצגת RTL.
4. עבור Outlook (opt-in): המגש מציג קודם דיאלוג אזהרה מפורט על חשיפת תוכן מייל ל-DOM של הפאנל. ברירת המחדל הממוקדת היא Cancel - לחיצת Enter בטעות לא מאשרת. רק אחרי OK מפורש המגש סוגר את Outlook ומפעיל אותו מחדש דרך `outlook-wrapper.bat` שכותב דגל opt-in זמני. ה-injector חוסם את Outlook אוטומטית ללא הדגל הזה. החיבור מתנתק אוטומטית אחרי 15 דקות, או ידנית עם **Disconnect Outlook only** בתפריט.

אפשר לחבר את כל ארבעת האפליקציות בו-זמנית, וה-injector יחיד יטפל בכולן.

סטטוס המגש במבט: ירוק, מחובר. אדום, מנותק או שגיאה. אפור, בהפעלה. בראש התפריט יש ארבע שורות סטטוס מנוטרלות לקריאה בלבד (Word, Excel, PowerPoint, Outlook - כל אחת במצב connected, not running, running without RTL, או error). מתחתיהן ארבעה פריטי Connect: **Connect Word**, **Connect Excel**, **Connect PowerPoint** ו-**Connect Outlook** (האחרון מציג דיאלוג opt-in לפני כל הפעלה). ולאחר מכן **Disconnect Outlook only** (פעיל רק כש-Outlook מחובר) ו-**Disconnect all** שסוגרת את Word/Excel/PowerPoint וה-injector אבל **לא** סוגרת את Outlook עצמו. בנוסף יש **Show diagnostic log**, **Check for updates...**, **Uninstall...** ו-**Exit**. אין יותר checkbox של Auto-enable - הוא הוסר בגרסה v0.1.4 משיקולי אבטחה (טריגר ל-EDR ארגוני).

לא משונים שיוכי קבצים, לא מתווספות רשומות לתפריט ההתחלה, ו-Word/Excel/PowerPoint/Outlook עצמם לא מתוקנים.

### מצב דיבאג (אופציונלי)

אם אתה רוצה לראות לוג חי בזמן שה-injector רץ (שימושי לדיווח תקלות או לבדיקת שינויים), הפעל בלחיצה כפולה את `start.bat` במקום להשתמש במגש. הוא פותח חלון לוג גלוי, סגור אותו כדי לעצור. לא נדרש לשימוש יומיומי.

## הסרה

שלוש דרכים שקולות, כולן מפעילות את אותו `uninstall.bat`:

- **Tray > Uninstall...** - קליק ימני על אייקון המגש, בחירה ב-Uninstall, אישור. המגש יוצא ומעביר את השליטה ל-`uninstall.bat`.
- **Windows Settings > Apps > Installed apps** - מוצאים "Claude for Word RTL Fix" ברשימה ולוחצים Uninstall. Windows מפעיל את אותו `uninstall.bat` דרך ערך `UninstallString` ברישום.
- **לחיצה כפולה על `uninstall.bat`** - הפעלה ישירה מתיקיית ההתקנה.

כל שלוש הדרכים מסירות: את קיצור ה-Startup, את הרישום ב-Apps and Features (`HKCU\...\Uninstall\ClaudeWordRTL`), את המגש וה-injector, ואת `node_modules`. אם משתנה הסביבה של Auto-enable תואם בדיוק לערך שנכתב ע"י הכלי - הוא גם מוסר. Word עצמו לא משתנה.

## עדכון לגרסה חדשה

שלושה שלבים:

1. הורידו את ה-ZIP החדש מ-[Releases](https://github.com/asaf-aizone/Claude-for-Office-RTL-fix/releases/latest) וחלצו מעל תיקיית ההתקנה הקיימת (החליפו קבצים כשמתבקש).
2. סגרו לחלוטין את כל אפליקציות Office הרלוונטיות (Word, Excel, PowerPoint), כולל תהליכי רקע דרך Task Manager במקרה הצורך.
3. הפעילו `install.bat` מחדש. הסקריפט עוצר את הטריי הישן דרך קובץ ה-PID לפני טעינת הקוד החדש, אז העדכון נכנס לתוקף מיד בלי צורך ב-logout.

לבדיקה שהגרסה החדשה אכן נטענה: `Check for updates...` בתפריט הטריי אמור להראות "You are on the latest version."

## אבחון, סטטוס ועדכונים

- **אייקון מגש** (tray) - אייקון קטן ליד השעון (ריבוע מעוגל עם האות **O** (Office) וחץ RTL לבנים, וצבע רקע שמשקף את המצב), מופעל אוטומטית בכניסה למערכת מתיקיית ה-Startup. ירוק, ה-injector מחובר לפאנל של Claude. אדום, מנותק או שדווחה שגיאה. אפור, בהפעלה. ראו את [האייקון במצב אדום (מנותק)](docs/images/tray-icon-red.png) ובמצב [ירוק (מחובר)](docs/images/tray-icon-green.png). לחיצה ימנית פותחת תפריט. בראש התפריט - ארבע שורות סטטוס לכל אחת מארבע האפליקציות (Word/Excel/PowerPoint/Outlook), ולאחר מכן: **Connect Word** / **Connect Excel** / **Connect PowerPoint** - מפעילים את האפליקציה הנבחרת דרך ה-wrapper אם היא סגורה, או אם היא כבר פתוחה "רגיל" (המקרה הנפוץ), שואלים את המשתמש, סוגרים אותה בצורה מנומסת ומפעילים אותה מחדש עם RTL - כולל פתיחה אוטומטית של הקבצים שהיו פתוחים. **Connect Outlook** - מסלול ייעודי opt-in עם דיאלוג אזהרה (ברירת מחדל Cancel) לפני כל הפעלה, כתיבת דגל opt-in זמני, ומסרב אם New Outlook (`olk.exe`) פתוח. **Disconnect Outlook only** - פעיל רק כש-Outlook מחובר, מנתק את ה-CDP של Outlook בלבד בלי לפגוע ב-Word/Excel/PowerPoint. **Disconnect all** - כפתור התאוששות כללי: עוצר timers של Connect באמצע, סוגר את Word/Excel/PowerPoint הפתוחים (מנומס + force כגיבוי), הורג את ה-injector, מנקה קבצי state ומבטל את דגל ה-opt-in של Outlook. **אינו סוגר את Outlook עצמו** (כי הוא opt-in וייתכן שהמשתמש פתח אותו רק לקריאת מייל); הריגת ה-injector ממילא מנתקת את ה-CDP. **Show diagnostic log** - פותח את `%TEMP%\claude-word-rtl.log` בעורך ברירת המחדל. **Check for updates...** - מריץ את `check-update.js` ומציג דיאלוג עם הסטטוס. אם יש גרסה חדשה, כפתור בלחיצה אחת פותח את דף ההורדה בדפדפן ברירת המחדל. **Uninstall...** - מציג אישור ואז מעביר את השליטה ל-`uninstall.bat` ויוצא. **Exit** - סוגר את ה-tray. רק מופע אחד של tray יכול לרוץ בכל רגע (נאכף ע"י mutex גלובלי), כדי שלא יראו שני אייקונים. בלי תלויות חדשות, הכל על בסיס `System.Windows.Forms.NotifyIcon` המובנה.
- **`doctor.bat`** - סקריפט אבחון שמריץ 19 בדיקות (החל מ-v0.3.0; היו 15 ב-v0.2.x). הבדיקות: Node, npm, תלויות, זיהוי התקנה לכל אחת מ-Word/Excel/PowerPoint, האם כל אפליקציה רצה כעת, פורטי CDP דינמיים שזוהו, יעדי Claude פעילים לפי אפליקציה, תהליך ה-injector, סטטוס מצרפי, סטטוס לפי אפליקציה (`apps.json`), תהליך ה-tray, רשומת Startup, רישום Apps and Features, הבדיקה הקריטית שאוסרת על משתנה הסביבה הישן `HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` לחזור, WebView2 runtime, וארבע בדיקות נוספות ייעודיות ל-Outlook: התקנה, תהליך רץ, target CDP, ומצב ב-`apps.json`. ארבע הבדיקות הללו כולן `INFO` כי Outlook הוא opt-in. הסקריפט כותב דוח לקובץ `doctor.log`. צרפו אותו כשמדווחים על תקלה.
- **`check-update.bat`** - פונה ל-GitHub releases API ומודיע אם יש גרסה חדשה יותר. אין תלויות npm, משתמש ב-`https` המובנה של Node. **איך בודקים אם יש גרסה חדשה?** הריצו `check-update.bat` או תפריט הטריי "Check for updates...". השוואה מול GitHub releases API, ללא תלויות חיצוניות.

## איך זה עובד (פסקה אחת)

לכל אחת מארבע האפליקציות יש wrapper משלה (`word-wrapper.bat`, `excel-wrapper.bat`, `powerpoint-wrapper.bat`, `outlook-wrapper.bat`). ה-wrapper של Outlook גם כותב דגל opt-in זמני (`%TEMP%\claude-office-rtl.outlook-optin`) שה-injector בודק לפני שמתחבר. ה-wrapper מפעיל את האפליקציה עם משתנה הסביבה `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0` - דגל של WebView2 המתועד אצל Microsoft שחושף את Chrome DevTools Protocol על פורט localhost דינמי. ערך 0 מאפשר ל-WebView2 לבחור פורט פנוי משלו לכל תהליך, מה שנדרש כשפועלות יותר מאפליקציית Office אחת בו-זמנית. המשתנה נשמר רק בהקשר של אותו תהליך wrapper ושל האפליקציה שהוא מפעיל - הוא לא מגיע ל-Teams, Outlook, Edge או כל WebView2 host אחר. `inject.js` יחיד מטפל בכל ארבע האפליקציות: בכל tick של 2 שניות הוא משתמש ב-`scripts/port-discovery.js` שמסקר את כל תהליכי `msedgewebview2.exe` דרך `tasklist`, ממפה אותם ל-LISTENING ports דרך `netstat`, ובודק כל פורט מועמד מול `/json/list` של CDP. עבור כל target הוא מזהה את האפליקציה (Word/Excel/Powerpoint/Outlook) דרך הפרמטר `_host_Info=` ב-URL של הפאנל, מתחבר ב-WebSocket, וקורא ל-`Runtime.evaluate` כדי להזריק אלמנט `<style>` ו-`MutationObserver`. אייקון המגש מתזמר את התהליך: בלחיצה על Connect של אחת מהאפליקציות הוא סוגר את האפליקציה הקיימת (אחרי שחילץ את הקבצים הפתוחים דרך COM ProgId המתאים) ומפעיל אותה מחדש דרך ה-wrapper המתאים עם אותם קבצים. כל הפעילות מקומית במחשב שלכם.

ראו [docs/security.md](docs/security.md) למודל האיומים.

## פתרון בעיות

**אבחון מהיר, לפני הטבלה: השתמשו ב-[Claude Code](https://claude.com/claude-code) ולא ב-Claude Chat.** Claude Code רץ מקומית ויכול לקרוא את `%TEMP%\claude-word-rtl.log`, את `%TEMP%\claude-office-rtl.apps.json` ואת `doctor.log`, ולהריץ `netstat` ו-`tasklist` כחלק מהאבחון; Chat לא רואה את הקבצים האלה. זרימה: להתקין את Claude Code, לפתוח session בתיקיית ההתקנה, לתאר את הבעיה בעברית. הוא יקרא את הלוגים ויציע תיקון.

- הקובץ `install.log` (נוצר בתיקיית ההתקנה) לוכד את הפלט המלא של הרצת ההתקנה האחרונה. צרפו אותו כשאתם מדווחים על תקלות.
- הפעילו את `cleanup.bat` אם תהליכי Node נשארים פעילים לאחר סגירת אפליקציית Office.
- אם הפאנל עדיין מוצג LTR, לחצו קליק ימני על אייקון המגש ובחרו ב-Connect המתאים לאפליקציה הפתוחה. אם המגש לא קיים, הפעילו בלחיצה כפולה את `scripts\start-tray.vbs`, או צאו מהחשבון וחזרו כדי שרשומת ה-Startup תיפעל.
- **הטריי נשאר אדום אחרי Connect, האפליקציה פתוחה, Node מותקן.** פותחים את `Show diagnostic log` מתפריט הטריי. הלוג מציג אילו פורטי CDP דינמיים נסרקו ואילו targets של Claude נמצאו. הריצו את `doctor.bat` כדי לראות את רשימת הפורטים והאפליקציות שזוהו (בדיקות 6 ו-7). אם רשימת הפורטים ריקה - האפליקציה לא נפתחה דרך Connect של ה-tray (פתיחה רגילה מהאייקון של Word/Excel/PowerPoint לא מפעילה את ה-debug port; חייבים לעבור דרך Connect או דרך ה-wrapper).

## מגבלות ידועות

- האפליקציה (Word/Excel/PowerPoint) חייבת להיות מופעלת דרך פעולת **Connect** של המגש (שמפעילה את ה-wrapper המתאים) או דרך ה-wrapper ישירות. פתיחה ישירה מהאייקון של האפליקציה לא מפעילה את פורט הדיבאג, אך המגש מזהה את המצב הזה ומציע להפעיל את האפליקציה מחדש דרך ה-wrapper עם הקבצים שהיו פתוחים.
- לא חל על טקסט ש-Claude כותב ישירות לגוף המסמך/חוברת העבודה/המצגת (זהו מסלול קוד נפרד מחוץ לחלונית ה-WebView2). הגדירו את גופן ברירת המחדל והסגנונות של האפליקציה לצורך זה.
- SmartScreen עלול להתריע בהפעלה ראשונה משום שהסקריפטים לא חתומים דיגיטלית.
- Device Guard / WDAC במחשבים מנוהלים עלול לחסום מתקינים וסקריפטים לא חתומים.
- **הכלי מסתמך על הזרקת CSS ו-JS ל-DOM של תוסף Claude, שעדיין בבטא.** עדכונים של Anthropic לתוסף עלולים לשנות את מבנה ה-DOM, שמות ה-classes או תבנית ה-URL ולשבור את הכלי ללא התראה. אם הכלי מפסיק לעבוד לאחר עדכון של Office, פתחו issue, מהדורה מתוקנת בדרך כלל היא שינוי של שורה אחת בבוררים.

## האם Anthropic יחסמו אותי?

לא צפוי. הכלי רק משנה את אופן התצוגה של הפאנל ב-DOM המקומי שלכם. הוא לא משנה מה אתם שולחים, מה Claude מחזיר, מגבלות שימוש או מנגנוני הגנה. השימוש שלכם ב-Claude נשאר כפוף לתנאי השירות ולמדיניות השימוש של Anthropic ללא קשר לכלי הזה. הכלי לא משנה את מה שנשלח ל-Claude או את מה ש-Claude מחזיר, אלא רק מעצב מחדש את הפלט שכבר עבר רינדור מקומית בדפדפן שלכם. אם תנאי Anthropic יתעדכנו אי פעם באופן שיגביל שינויים בצד הלקוח, יש לציית לתנאים שלהם ולא לכלי הזה.

## כתב ויתור

זהו כלי נגישות עצמאי וקוד פתוח. הוא **אינו** מזוהה, מאושר או מחובר ל-Anthropic או ל-Microsoft.

- "Claude" הוא סימן מסחרי של Anthropic, PBC.
- "Microsoft", "Word", "Excel" ו-"PowerPoint" הם סימנים מסחריים של Microsoft Corporation.
- הפרויקט הזה לא מפיץ, לא משנה ולא מכיל קוד קנייני משתיהן.

בשימוש בכלי זה, אתם מאשרים שאתם אחראים להבטיח שהשימוש שלכם ב-Claude וב-Microsoft Office (Word, Excel, PowerPoint) תואם את תנאי השירות שלהם לסוג החשבון שברשותכם. ראה את [Consumer Terms](https://www.anthropic.com/legal/consumer-terms), [Commercial Terms](https://www.anthropic.com/legal/commercial-terms) ו-[Acceptable Use Policy](https://www.anthropic.com/legal/aup) של Anthropic. ראה את [Office Add-ins Privacy and Security](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/privacy-and-security) של Microsoft.

## תודות

- נוצר על ידי **Asaf Abramzon** - [LinkedIn](https://www.linkedin.com/in/asaf-abramzon-7a2b61180/) · [GitHub](https://github.com/asaf-aizone).
- [`chrome-remote-interface`](https://github.com/cyrus-and/chrome-remote-interface) - CDP client.

## פרויקטים קשורים בקהילת RTL-for-AI-tools

מספר מפתחים עצמאיים בנו פתרונות RTL בשכבות שונות של סביבת ה-AI tooling. כל פרויקט מטפל בסביבה שונה:

- **[Adaptive-RTL-Extension](https://github.com/Lidor-Mashiach/Adaptive-RTL-Extension)** מאת לידור משיח - הרחבת דפדפן גנרית עם click-to-select ל-RTL בכל אתר, כולל ממשקי צ'אט של LLMים.
- **[Claude.ai RTL Support](https://chromewebstore.google.com/detail/claude-ai-rtl-support/lkopcjdmfmffphbomfhecalbojiaeape)** - הרחבת Chrome ייעודית ל-Claude.ai. קלה יותר מהגנרית אם צריך RTL רק על ממשק הווב.
- **[rtl-for-vs-code-agents](https://github.com/GuyRonnen/rtl-for-vs-code-agents)** מאת גיא רונן - הרחבה ל-VS Code עבור Claude Code, Cursor, Antigravity ו-Gemini Code Assist בשכבת ה-webview.
- **[Claude Code RTL Support](https://open-vsx.org/extension/yechielby/claude-code-rtl)** מאת יחיאל בר-יהודה - תוסף ייעודי לפאנל Claude Code בתוך VS Code/Cursor/Antigravity. 6,700+ התקנות.
- **Claude for Office RTL Fix** *(הפרויקט הזה)* - תוסף Claude ל-Word, Excel, PowerPoint ו-Outlook ב-Microsoft Office desktop.
- **[kivun-terminal-wsl](https://github.com/noambrand/kivun-terminal-wsl)** מאת נועם ברנד - תיקון RTL בשכבת הטרמינל של Claude Code על Windows/Linux.

הסביבות נפרדות זו מזו - בחרו את הפתרון שמתאים למקום שבו אתם נתקלים בבעיית ה-BiDi.

## רישיון

Apache License 2.0, ראה [LICENSE](LICENSE).

## תרומה

Issues ו-pull requests מתקבלים בברכה. אם מצאת באג תצוגה (בורר שדלף, יישור שגוי במקום מסוים), פתח issue עם צילום מסך וצרף את `install.log` שלך.

</div>
