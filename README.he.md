<div dir="rtl">

# Claude for Word RTL Fix

[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue)](LICENSE)
[![Node](https://img.shields.io/badge/node-%E2%89%A516-brightgreen)](https://nodejs.org/)
[![No Telemetry](https://img.shields.io/badge/telemetry-none-success)](#פרטיות)
[![Local Only](https://img.shields.io/badge/network-localhost%20only-success)](#הערת-אבטחה)

**גרסה מלאה (דו-לשונית, עם צילומי מסך ותמונות): [README.md](README.md)** · **יומן שינויים: [CHANGELOG.md](CHANGELOG.md)**

תיקון CSS וטיפוגרפיה בצד הלקוח לתצוגה העברית בתוסף Claude ל-Microsoft Word.

תוסף Claude הרשמי ל-Word מציג כיום טקסט עברי משמאל לימין, עם סימני רשימה ופיסוק בצד הלא נכון. הכלי הזה מתחבר לחלונית WebView2 של התוסף באמצעות Chrome DevTools Protocol הסטנדרטי, ומזריק גיליון סגנונות וכן MutationObserver קטן כדי לתקן את התצוגה.

הכל פועל מקומית במחשב שלך. שום דבר לא נשלח ברשת.

> **Windows בלבד.** הכלי לא עובד על macOS או Linux. תוסף Claude ל-Word מבוסס על WebView2 של מיקרוסופט, שקיים רק ב-Windows. ל-Word ל-Mac יש runtime אחר (WKWebView) שלא חושף את אותו debugging interface, וכל שכבת ההפעלה (bat, vbs, PowerShell, Registry, Startup folder) לא רלוונטית שם. אם אתם על Mac, אין port מ-Word.

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
- לא שולח טלמטריה או תעבורת רשת משלו
- לא שומר אישורים, תוכן שיחות או כל מידע אחר
- לא משנה שיוכי קבצים של Word
- לא יוצר משימות מתוזמנות או שירותי רקע
- לא משנה את `Normal.dotm` או כל תבנית אחרת של Word
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

בזמן ש-Word פועל דרך הכלי הזה, WebView2 פותח פורט דיבאג על `localhost:9222`. משמעות הדבר שכל תהליך מקומי אחר במחשב שלך יכול להתחבר ל-DOM של חלונית Claude (לקרוא טיוטות, עוגיות סשן וכדומה). הפורט הוא localhost בלבד, לא חשוף לרשת, אבל הוא לא דורש אימות.

המלצות:

- סגור את Word כשאתה לא משתמש ב-Claude באופן פעיל.
- אל תריץ את הכלי על מחשבים משותפים או מחשבים עם תוכנות לא מהימנות.
- במחשבים מנוהלים ארגונית (EDR, DLP), בדוק תחילה עם ה-IT.

ראה [SECURITY.md](SECURITY.md) למודל האיומים המלא ולתהליך דיווח פגיעויות.

## דרישות

- **Windows 10 או 11 בלבד.** macOS ו-Linux לא נתמכים (ראו למעלה).
- Microsoft Word (דסקטופ), עם התוסף Claude מותקן
- [Node.js](https://nodejs.org/) 16 או חדש יותר (מותקן ונמצא ב-PATH)

## התקנה

1. שכפל את המאגר או הורד כ-ZIP.
2. סגור את Microsoft Word אם הוא פתוח (המתקין יבדוק ויתריע).
3. הפעל בלחיצה כפולה על `install.bat`.
   - בהפעלה הראשונה הוא מתקין את `chrome-remote-interface` דרך npm.
   - לוג התקנה מלא נכתב לקובץ `install.log` לצד הסקריפט.
   - Windows SmartScreen עלול להתריע. לחץ על "מידע נוסף" ואז "הפעל בכל זאת" אם אתה סומך על המקור.
   - לא נדרשות הרשאות מנהל. המתקין יוצר קיצור יחיד בתיקיית Startup (למגש) ומפתח רישום יחיד תחת `HKCU\...\Uninstall\ClaudeWordRTL` כדי שהכלי יופיע ב-Windows Settings > Apps.

### איך משתמשים

אייקון המגש הוא נקודת הכניסה היחידה:

1. פתח את Word כרגיל (דרך אייקון Word, קיצור דרך, מסמך, כל מה שאתה רגיל להשתמש בו).
2. לחץ קליק ימני על אייקון המגש ליד השעון ובחר **Connect**.
3. המגש סוגר בצורה מנומסת את Word, מפעיל אותו מחדש דרך ה-wrapper עם דגל הדיבאג, ופותח מחדש את המסמכים שהיו פתוחים. עברית בפאנל של Claude עכשיו מוצגת RTL.

סטטוס המגש במבט: ירוק, מחובר. אדום, מנותק או שגיאה. אפור, בהפעלה. בתפריט המגש יש גם **Disconnect (close Claude for Word RTL Fix)**, **Show diagnostic log**, **Check for updates...**, **Uninstall...** ו-**Exit**.

לא משונים שיוכי קבצים, לא מתווספות רשומות לתפריט ההתחלה, ו-Word עצמו לא מתוקן.

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

1. הורידו את ה-ZIP החדש מ-[Releases](https://github.com/asaf-aizone/Claude-for-word-RTL-fix/releases/latest) וחלצו מעל תיקיית ההתקנה הקיימת (החליפו קבצים כשמתבקש).
2. סגרו את Word לחלוטין (גם תהליכי רקע דרך Task Manager במקרה הצורך).
3. הפעילו `install.bat` מחדש. הסקריפט עוצר את הטריי הישן דרך קובץ ה-PID לפני טעינת הקוד החדש, אז העדכון נכנס לתוקף מיד בלי צורך ב-logout.

לבדיקה שהגרסה החדשה אכן נטענה: `Check for updates...` בתפריט הטריי אמור להראות "You are on the latest version."

## אבחון, סטטוס ועדכונים

- **אייקון מגש** (tray) - אייקון קטן ליד השעון (ריבוע מעוגל עם האות **W** וחץ RTL לבנים, וצבע רקע שמשקף את המצב), מופעל אוטומטית בכניסה למערכת מתיקיית ה-Startup. ירוק, ה-injector מחובר לפאנל של Claude. אדום, מנותק או שדווחה שגיאה. אפור, בהפעלה. ראה את [האייקון במצב אדום (מנותק)](docs/images/tray-icon-red.png) ובמצב [ירוק (מחובר)](docs/images/tray-icon-green.png). לחיצה ימנית פותחת תפריט. **Connect (relaunch Claude for Word RTL Fix)** - מפעיל את Word דרך ה-wrapper אם הוא סגור, או אם Word כבר פתוח "רגיל" (המקרה הנפוץ), שואל את המשתמש, סוגר אותו בצורה מנומסת ומפעיל אותו מחדש עם RTL - כולל פתיחה אוטומטית של המסמכים שהיו פתוחים. **Disconnect (close Claude for Word RTL Fix)** - כפתור התאוששות כללי: עוצר timers של Connect באמצע, סוגר את Word (מנומס + force כגיבוי), הורג את ה-injector, מנקה קבצי state. **Show diagnostic log** - פותח את `%TEMP%\claude-word-rtl.log` בעורך ברירת המחדל. **Check for updates...** - מריץ את `check-update.js` ומציג דיאלוג עם הסטטוס. אם יש גרסה חדשה, כפתור בלחיצה אחת פותח את דף ההורדה בדפדפן ברירת המחדל. **Uninstall...** - מציג אישור ואז מעביר את השליטה ל-`uninstall.bat` ויוצא. **Exit** - סוגר את ה-tray. רק מופע אחד של tray יכול לרוץ בכל רגע (נאכף ע"י mutex גלובלי), כדי שלא יראו שני אייקונים. בלי תלויות חדשות, הכל על בסיס `System.Windows.Forms.NotifyIcon` המובנה.
- **`doctor.bat`** - סקריפט אבחון שמריץ 12 בדיקות (Node, npm, תלויות, התקנת Word, פורט 9222, תהליך ה-injector, רשומת Startup, תהליך ה-tray, WebView2 runtime, גרסת Office, רישום ב-Apps and Features) וכותב דוח לקובץ `doctor.log`. צרף אותו כשמדווחים על תקלה.
- **`check-update.bat`** - פונה ל-GitHub releases API ומודיע אם יש גרסה חדשה יותר. אין תלויות npm, משתמש ב-`https` המובנה של Node. **איך בודקים אם יש גרסה חדשה?** הריצו `check-update.bat` או תפריט הטריי "Check for updates...". השוואה מול GitHub releases API, ללא תלויות חיצוניות.

## איך זה עובד (פסקה אחת)

`word-wrapper.bat` מפעיל את Word עם משתנה הסביבה `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`. זהו דגל של WebView2 המתועד אצל Microsoft שחושף את Chrome DevTools Protocol על `localhost:9222`. `inject.js` מתחבר לפורט הזה, מאתר את יעד ה-WebView שכתובתו תואמת ל-`claude.ai`, וקורא ל-`Runtime.evaluate` כדי להזריק אלמנט `<style>` ו-`MutationObserver`. לולאת פולינג של שתי שניות מזריקה מחדש אם החלונית נטענת מחדש. אייקון המגש מתזמר את התהליך: בלחיצה על **Connect** הוא סוגר את Word הקיים (אחרי שחילץ את המסמכים הפתוחים דרך COM) ומפעיל אותו מחדש דרך ה-wrapper עם אותם מסמכים. כל הפעילות מקומית במחשב שלך.

ראה [docs/security.md](docs/security.md) למודל האיומים.

## פתרון בעיות

**אבחון מהיר, לפני הטבלה: השתמשו ב-[Claude Code](https://claude.com/claude-code) ולא ב-Claude Chat.** Claude Code רץ מקומית ויכול לקרוא את `%TEMP%\claude-word-rtl.log` ו-`doctor.log` ולהריץ `netstat` כחלק מהאבחון; Chat לא רואה את הקבצים האלה. זרימה: להתקין את Claude Code, לפתוח session בתיקיית ההתקנה, לתאר את הבעיה בעברית. הוא יקרא את הלוגים ויציע תיקון.

- הקובץ `install.log` (נוצר בתיקיית ההתקנה) לוכד את הפלט המלא של הרצת ההתקנה האחרונה. צרף אותו כשאתה מדווח על תקלות.
- הפעל את `cleanup.bat` אם תהליכי Node נשארים פעילים לאחר סגירת Word.
- אם הפאנל עדיין מוצג LTR, לחץ קליק ימני על אייקון המגש ובחר **Connect**. אם המגש לא קיים, הפעל בלחיצה כפולה את `scripts\start-tray.vbs`, או צא מהחשבון וחזור כדי שרשומת ה-Startup תיפעל.
- **הטריי נשאר אדום למרות ש-Auto-enable דלוק, Word פתוח, ו-Node מותקן.** פורט 9222 אולי תפוס בידי אפליקציה אחרת. בודקים ב-cmd: `netstat -ano | findstr :9222`. אם רואים שורה עם PID שאינו של `WINWORD.EXE`, אפליקציה אחרת יושבת על הפורט. אשמים מוכרים: Google Drive File Stream, אפליקציות מבוססות Electron שמריצות עם `--remote-debugging-port=9222`, או WebView2 SDK tools. סוגרים את האפליקציה שתופסת את הפורט (לפי ה-PID וה-process name ב-Task Manager) ואז פותחים את Word מחדש דרך הטריי (Connect). מגרסה 0.1.3 ואילך ה-injector מטפל במקרה שבו גם Drive וגם Word יושבים על 9222 בו-זמנית בגלל IPv4/IPv6 split, אבל אם אף אחד לא מציע panel של Claude, אין מה לתפוס. `doctor.bat` של גרסה 0.1.3 מציג את זה בבירור בשתי הבדיקות החדשות.

## מגבלות ידועות

- Word חייב להיות מופעל דרך פעולת **Connect** של המגש (שמפעילה את ה-wrapper). פתיחת Word ישירות מהאייקון של עצמו לא מפעילה את פורט הדיבאג, אך המגש מזהה את המצב הזה ומציע להפעיל את Word מחדש דרך ה-wrapper עם המסמכים שהיו פתוחים.
- לא חל על טקסט ש-Claude כותב ישירות לגוף מסמך ה-Word (זהו מסלול קוד נפרד מחוץ לחלונית ה-WebView2). הגדר את גופן ברירת המחדל והסגנונות של Word לצורך זה.
- SmartScreen עלול להתריע בהפעלה ראשונה משום שהסקריפטים לא חתומים דיגיטלית.
- Device Guard / WDAC במחשבים מנוהלים עלול לחסום מתקינים וסקריפטים לא חתומים.
- **הכלי מסתמך על הזרקת CSS ו-JS ל-DOM של תוסף Claude, שעדיין בבטא.** עדכונים של Anthropic לתוסף עלולים לשנות את מבנה ה-DOM, שמות ה-classes או תבנית ה-URL ולשבור את הכלי ללא התראה. אם הכלי מפסיק לעבוד לאחר עדכון של Word, פתח issue, מהדורה מתוקנת בדרך כלל היא שינוי של שורה אחת בבוררים.

## כתב ויתור

זהו כלי נגישות עצמאי וקוד פתוח. הוא **אינו** מזוהה, מאושר או מחובר ל-Anthropic או ל-Microsoft.

- "Claude" הוא סימן מסחרי של Anthropic, PBC.
- "Microsoft" ו-"Word" הם סימנים מסחריים של Microsoft Corporation.
- הפרויקט הזה לא מפיץ, לא משנה ולא מכיל קוד קנייני משתיהן.

בשימוש בכלי זה, אתה מאשר שאתה אחראי להבטיח שהשימוש שלך ב-Claude וב-Microsoft Word תואם את תנאי השירות שלהם לסוג החשבון שברשותך. ראה את [Consumer Terms](https://www.anthropic.com/legal/consumer-terms), [Commercial Terms](https://www.anthropic.com/legal/commercial-terms) ו-[Acceptable Use Policy](https://www.anthropic.com/legal/aup) של Anthropic. ראה את [Office Add-ins Privacy and Security](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/privacy-and-security) של Microsoft.

## תודות

- נוצר על ידי **Asaf Abramzon** - [LinkedIn](https://www.linkedin.com/in/asaf-abramzon-7a2b61180/) · [GitHub](https://github.com/asaf-aizone).
- [`chrome-remote-interface`](https://github.com/cyrus-and/chrome-remote-interface) - CDP client.

## רישיון

Apache License 2.0, ראה [LICENSE](LICENSE).

## תרומה

Issues ו-pull requests מתקבלים בברכה. אם מצאת באג תצוגה (בורר שדלף, יישור שגוי במקום מסוים), פתח issue עם צילום מסך וצרף את `install.log` שלך.

</div>
