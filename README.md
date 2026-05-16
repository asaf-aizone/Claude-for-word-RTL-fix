<div dir="rtl">

<h1>Claude for Office RTL Fix (Word, Excel, PowerPoint, Outlook)</h1>

<p>
תיקון RTL לפאנל של Claude ב-Microsoft Word, Excel, PowerPoint ו-Outlook. מקומי, בלי טלמטריה, Apache 2.0.<br>
<em>לכל מי שכותב עברית באפליקציות Office עם תוסף Claude. Windows 10/11.</em><br>
<strong>חדש ב-v0.3.0:</strong> תמיכה ב-Outlook הקלאסי. מודל אבטחה מוקשח (opt-in מפורש כל הפעלה, ניתוק אוטומטי אחרי 15 דקות, אזהרה לפני הפעלה) - ראו את הסקציה <a href="#outlook-section">Outlook (opt-in)</a>.
</p>

<p>
<strong>אינו תוסף רשמי של Anthropic או של Microsoft.</strong> כלי open-source עצמאי. שם המאגר ב-GitHub עודכן ב-v0.2.1 ל-<code>Claude-for-Office-RTL-fix</code> (היה <code>Claude-for-word-RTL-fix</code>). GitHub שומר על redirect קבוע מהשם הישן, ולכן clones, bookmarks ו-clone URLs מגרסאות v0.1.x ממשיכים לעבוד.
</p>

<blockquote>
<p>
<strong>Windows בלבד.</strong> הכלי לא עובד על macOS או Linux. תוסף Claude ל-Office
מבוסס על WebView2 של מיקרוסופט, שקיים רק ב-Windows. ל-Office ל-Mac יש runtime
אחר (WKWebView) שלא חושף את אותו debugging interface, וכל שכבת ההפעלה (bat, vbs,
PowerShell, Registry, Startup folder) לא רלוונטית שם. אם אתם על Mac, אין port מ-Office.
</p>
</blockquote>

<blockquote>
<p>
<strong>אזהרה למחשבים מנוהלי-ארגון.</strong> הכלי מתחבר ל-Microsoft Word דרך
Chrome DevTools Protocol ומזריק JavaScript לתוך WebView2, ומפעיל את עצמו דרך
VBS hidden launcher ו-PowerShell. הצירוף הזה דומה מבחינה מבנית לטכניקות שגונבי-מידע
משתמשים בהן, ולכן מערכות EDR ארגוניות (Microsoft Defender for Endpoint, CrowdStrike
Falcon, SentinelOne, Sophos) עלולות לזהות את ההתקנה כפעילות חשודה ולנתק את המכונה
מהרשת (host isolation) באופן אוטומטי. <strong>אין להתקין על מחשב מנוהל-ארגון בלי
אישור מקדים מצוות אבטחת המידע</strong> ובלי הוספת ה-hash וה-path של הקבצים ל-allowlist.
המחבר אינו אחראי לתגובות מערכות אבטחה ארגוניות.
</p>
</blockquote>

<p>
  <a href="#install"><img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue" alt="Platform"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue" alt="License"></a>
  <a href="https://nodejs.org/"><img src="https://img.shields.io/badge/node-%E2%89%A516-brightgreen" alt="Node 16+"></a>
  <img src="https://img.shields.io/badge/telemetry-none-success" alt="No Telemetry">
  <img src="https://img.shields.io/badge/network-localhost%20only-success" alt="Local Only">
</p>

<hr>

<h2>מה זה עושה?</h2>

<p>
הפאנל של Claude בתוך אפליקציות Office (Word, Excel, PowerPoint) לא תומך
ב-RTL. עברית יוצאת הפוכה, bullets בצד הלא נכון, פיסוק נופל איפה שלא צריך.
הכלי מתחבר ל-WebView2 של הפאנל בכל אחת משלוש האפליקציות דרך Chrome
DevTools Protocol, מזריק CSS ו-MutationObserver, והפאנל עובר ל-RTL תקין.
</p>

<p>
המודל של Claude, ה-API של Anthropic, ואפליקציות Office עצמן לא נגועים. רק ה-DOM
המקומי של הפאנל, ורק כל עוד הפאנל פתוח.
</p>

<hr>

<h2>פיצ'רים</h2>

<h3>הפאנל עובר ל-RTL</h3>
<p>
כיוון טקסט, יישור, bullets, טבלאות. בלוקי קוד (<code>&lt;pre&gt;</code>, <code>&lt;code&gt;</code>)
נשארים LTR כך שקוד לא מתעוות.
</p>
<p>
  <img src="docs/images/before.png" alt="הפאנל לפני - em-dash, לא מיושר מלא" width="420">
  <img src="docs/images/after.png" alt="הפאנל אחרי - RTL תקין, hyphen במקום em-dash, scrollbar שמאל" width="420">
  <br>
  <em>מימין - הפאנל לפני. משמאל - אחרי RTL. אותה שאלה, אותה תשובה.</em>
</p>

<h3>התערבות מינימלית ב-Word</h3>
<p>
הכלי לא משנה file associations, לא נוגע ב-<code>Normal.dotm</code>, לא
מוסיף טמפלטים או תוספים, ולא יוצר services. ההתקנה כן יוצרת שני דברים
ברמת המשתמש: קיצור ב-Startup folder שמעלה את ה-tray בלוגין (ניתן
למחוק ידנית מ-<code>shell:startup</code>), ומפתח רישום תחת
<code>HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL</code>
כדי שהכלי יופיע ב-Windows Settings &gt; Apps. שניהם נמחקים ע"י
<code>uninstall.bat</code>.
</p>

<h3>אייקון Tray ליד השעון</h3>
<p>
אייקון קטן ליד השעון: ריבוע מעוגל עם האות <strong>O</strong> (Office) וחץ RTL
לבנים, וצבע רקע שמשקף את המצב. אפור = בטעינה, אדום = מנותק, ירוק =
RTL פעיל באחת או יותר משלוש אפליקציות Office. קליק ימני פותח תפריט קצר. אין כניסה ב-Start Menu, אין שינוי
של קיצורי מערכת.
</p>
<p>
  <img src="docs/images/tray-icon-red.png" alt="אייקון אדום - מנותק" width="80">
  &nbsp;&nbsp;
  <img src="docs/images/tray-icon-green.png" alt="אייקון ירוק - מחובר" width="80">
  <br>
  <em>מימין: אייקון אדום (מנותק או בהפעלה). משמאל: אייקון ירוק (מחובר, RTL פעיל בפאנל).
  צילומי מסך אמיתיים מה-tray של Windows.</em>
</p>

<table dir="rtl">
  <thead>
    <tr><th>פריט בתפריט</th><th>מתי זמין</th><th>מה עושה</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>Word: ... / Excel: ... / PowerPoint: ... / Outlook: ...</td>
      <td>תמיד (ארבע שורות מנוטרלות בראש התפריט)</td>
      <td>תוויות סטטוס לקריאה בלבד, אחת לכל אפליקציה. כל שורה יכולה להציג <code>connected</code>, <code>not running</code>, <code>running without RTL</code>, או <code>error</code>. מתעדכנות כל 2 שניות מתוך <code>%TEMP%\claude-office-rtl.apps.json</code></td>
    </tr>
    <tr>
      <td>Connect Word</td>
      <td>תמיד</td>
      <td>מעלה את Word מחדש דרך <code>word-wrapper.bat</code> עם debug-port פתוח (פורט דינמי). אם Word רץ - מונה את המסמכים הפתוחים דרך COM, מבקש אישור, מריץ מחדש עם אותם מסמכים</td>
    </tr>
    <tr>
      <td>Connect Excel</td>
      <td>תמיד</td>
      <td>אותו זרם ל-Excel דרך <code>excel-wrapper.bat</code>. מונה דרך <code>Workbooks</code> במקום <code>Documents</code></td>
    </tr>
    <tr>
      <td>Connect PowerPoint</td>
      <td>תמיד</td>
      <td>אותו זרם ל-PowerPoint דרך <code>powerpoint-wrapper.bat</code>. מונה דרך <code>Presentations</code></td>
    </tr>
    <tr>
      <td>Connect Outlook</td>
      <td>תמיד (פריט opt-in)</td>
      <td>מציג דיאלוג אזהרה ייעודי (ברירת המחדל היא Cancel) שמסביר שבזמן Connect לתוכן מייל פעיל יש חשיפה ל-DOM של הפאנל. אם המשתמש מאשר, מפעיל את <code>outlook-wrapper.bat</code> שכותב דגל opt-in ייעודי (<code>%TEMP%\claude-office-rtl.outlook-optin</code>) ומפעיל את Outlook הקלאסי. ה-injector חוסם את Outlook אוטומטית בלי הדגל הזה, ולכן השלב הזה הוא תנאי הכרחי. מסרב לרוץ אם New Outlook (<code>olk.exe</code>) פתוח</td>
    </tr>
    <tr>
      <td>Disconnect Outlook only</td>
      <td>רק כש-Outlook במצב <code>connected</code></td>
      <td>מנתק את ה-CDP של Outlook בלבד בלי לסגור את Outlook עצמו ובלי לפגוע ב-Word/Excel/PowerPoint. המימוש כותב קובץ בקשה (<code>claude-office-rtl.disconnect-outlook.request</code>) שה-injector קורא בכל tick, סוגר את הלקוחות של Outlook, מבטל את timer ההתנתקות ומבטל את דגל ה-opt-in. שימושי בסוף סשן עבודה במייל כשעדיין רוצים את Word/Excel/PowerPoint מחוברים</td>
    </tr>
    <tr>
      <td>Disconnect all</td>
      <td>תמיד</td>
      <td>כפתור התאוששות אוניברסלי. עוצר timers של Connect באוויר, סוגר את Word/Excel/PowerPoint הפתוחים (graceful + force כגיבוי), הורג את ה-injector, מנקה את כל ה-state files ומבטל את דגל ה-opt-in של Outlook. <strong>אינו סוגר את Outlook עצמו</strong> (כי הוא opt-in וייתכן שהמשתמש פתח אותו לקריאת מייל בלי קשר ל-RTL); הריגת ה-injector ממילא מנתקת את ה-CDP. לחיצה כאן תמיד מחזירה למצב אפס</td>
    </tr>
    <tr>
      <td>Show diagnostic log</td>
      <td>תמיד</td>
      <td>פותח את <code>%TEMP%\claude-word-rtl.log</code> ב-editor של ברירת המחדל. שימושי כשמשהו לא עובד</td>
    </tr>
    <tr>
      <td>Check for updates...</td>
      <td>תמיד</td>
      <td>מפעיל את <code>check-update.js</code> ומציג דיאלוג עם הסטטוס. אם יש גרסה חדשה, כפתור בלחיצה אחת פותח את דף ההורדה בדפדפן</td>
    </tr>
    <tr>
      <td>Uninstall...</td>
      <td>תמיד</td>
      <td>מפעיל את <code>uninstall.bat</code> עם אישור</td>
    </tr>
    <tr>
      <td>Exit</td>
      <td>תמיד</td>
      <td>עוצר את האייקון</td>
    </tr>
  </tbody>
</table>

<p>
  <img src="docs/images/tray-menu.png" alt="תפריט קליק-ימני של ה-tray" width="360">
</p>

<h3>אקטיבציה דרך Connect</h3>

<p>
החל מ-v0.1.4 ה-Auto-enable הקבוע הוסר לחלוטין. הסיבה: הוא דרש כתיבה
של <code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code> ברמת
<code>HKCU\Environment</code>, משתנה שנקרא על ידי <em>כל</em> תהליך WebView2
שרץ תחת אותו user (Teams, Outlook, Edge WebView, OneDrive UI). מערכות
EDR ארגוניות (Microsoft Defender for Endpoint, CrowdStrike Falcon,
SentinelOne, Sophos) מסמנות שינויים במשתנה הזה כסיגנל לגניבת אישורים,
ובמקרה אחד בשטח זה גרם להן לבצע host isolation אוטומטי על מכונה מנוהלת.
מ-v0.2.0 התמונה זהה: אין משתנה סביבה קבוע, אין checkbox, אין prompt
במתקין.
</p>
<p>
במקום זאת, האקטיבציה כולה עוברת דרך <strong>Connect</strong> בתפריט
הטריי. כל לחיצה על Connect Word/Excel/PowerPoint קוראת ל-wrapper
הייעודי של אותה אפליקציה
(<code>word-wrapper.bat</code>, <code>excel-wrapper.bat</code>,
<code>powerpoint-wrapper.bat</code>), שמגדיר את משתנה ה-WebView2 רק
ב-process scope של עצמו וירש על ידי האפליקציה שהוא מפעיל. תהליכים
אחרים של WebView2 על המחשב לא רואים את המשתנה. כך השמירה על הפונקציונליות
המקורית של Auto-enable (RTL בכל פתיחה רלוונטית) מתבצעת בלי כתיבה
ל-Registry שמסמנת אותנו ל-EDR.
</p>

<p>
ב-uninstall, אם משתנה ה-Auto-enable הישן עדיין שם מגרסה v0.1.x, הוא
מנוקה אוטומטית - אבל רק אם הערך של המשתנה תואם בדיוק לאחד מהערכים שלנו
(<code>=9222</code> או <code>=0</code>). ערך שמשתמש הוסיף ידנית לא נמחק.
</p>

<h3 id="outlook-section">Outlook (opt-in, חדש ב-v0.3.0)</h3>
<p>
החל מ-v0.3.0 הכלי תומך גם ב-Outlook הקלאסי. עם זאת, Outlook מטופל אחרת
משלוש האפליקציות האחרות, ומשתי סיבות:
</p>
<ol>
  <li>כשמשתמש מבקש מ-Claude "Summarize this email" או "Draft a reply", תוכן המייל הפעיל הופך לחלק מה-DOM של הפאנל. ה-CDP attach שמאפשר ל-injector להזריק CSS הוא אותו attach שמאפשר לכל תהליך מקומי לקרוא את אותו DOM. מילים אחרות: כל עוד החיבור פתוח, תוכן המייל הפעיל חשוף לכל תוכנה אחרת שרצה תחת אותו user.</li>
  <li>תוכן מייל יכול לכלול סיסמאות זמניות, קודי MFA, מסמכים משפטיים, מזהי tenant - דברים שהחשיפה אליהם איכותית מסוכנת יותר מאשר חשיפה לתוכן מסמך Word שמשתמש העתיק במכוון.</li>
</ol>
<p>
לכן Outlook מוגן בארבע שכבות שלא פעילות עבור Word/Excel/PowerPoint:
</p>
<ul>
  <li><strong>Opt-in מפורש כל הפעלה.</strong> ה-injector חוסם את Outlook אוטומטית. לחיצה על <strong>Connect Outlook</strong> מציגה דיאלוג אזהרה מפורט (ברירת המחדל הממוקדת היא Cancel - לחיצת Enter בטעות לא מאשרת). רק אם המשתמש לוחץ OK מפורשות, ה-wrapper כותב דגל opt-in זמני שמאפשר ל-injector להתחבר. הדגל נמחק אוטומטית בכל הפעלה של ה-injector, ולכן הסכמה משיחה קודמת לא נשמרת.</li>
  <li><strong>אין auto-launch.</strong> ה-tray מפעיל מחדש את ה-injector אוטומטית רק אם Word/Excel/PowerPoint חיים בלי injector. Outlook לבדו אינו מצדיק הפעלה - אם המשתמש פתח את Outlook לקריאת מייל בלי כוונה ל-Claude, ה-injector נשאר מכובה.</li>
  <li><strong>ניתוק אוטומטי אחרי 15 דקות.</strong> גם אם המשתמש שכח לנתק ידנית, ה-injector מנתק את Outlook אחרי 15 דקות של חיבור רציף ומבטל את דגל ה-opt-in. כדי לחדש - יש ללחוץ Connect Outlook שוב. ערך זה ניתן להתאמה ב-<code>OUTLOOK_AUTO_DISCONNECT_MIN</code> בתוך <code>scripts/inject.js</code>.</li>
  <li><strong>Redaction של ה-URL בלוג.</strong> ה-URL של תוסף Office מכיל פרמטר <code>et=</code> שהוא base64 של מטא-דאטה של ה-tenant (account id, tenant id, expiry). עבור Outlook בלבד, ה-injector חותך מהלוג את כל הפרמטרים מלבד <code>_host_Info=</code>. ה-URL של Word/Excel/PowerPoint נשאר ללא שינוי לתאימות אבחון.</li>
</ul>
<p>
בנוסף, פריט תפריט ייעודי <strong>Disconnect Outlook only</strong> מאפשר לסגור
רק את חיבור ה-CDP של Outlook בלי לפגוע ב-Word/Excel/PowerPoint שעדיין מחוברים.
</p>
<p>
<strong>New Outlook (<code>olk.exe</code>) אינו נתמך.</strong> ההסתעפות אליו דחויה מאז M0,
ו-<code>outlook-wrapper.bat</code> וגם Connect Outlook מסרבים לרוץ אם הוא פתוח.
</p>
<p>
המודל המלא: <a href="docs/security.md#outlook-specific-risks-and-mitigations">סקציית Outlook-specific risks and mitigations</a> ב-<code>docs/security.md</code>, ולמי שרוצה להבין את ה-design - <a href="docs/OUTLOOK-EXPANSION-PLAN.md"><code>docs/OUTLOOK-EXPANSION-PLAN.md</code></a>.
</p>

<h3>ניקוי טיפוגרפי</h3>
<p>
em-dash (—) ו-en-dash (–) מוחלפים ב-hyphen (-). חצים (→ ← ↔ ⇒ ⇐)
מוחלפים בפסיק. שדות קלט לא נוגעים בהם, כך שמה שאתה מקליד נשאר כמו שהוא.
</p>

<h3>סטטוס שמשתדל לא להטעות</h3>
<p>
אם ה-injector מת בלי cleanup, ה-tray מזהה status מת ונהפך לאדום, כדי
להפחית מצבים שבהם האייקון ירוק בזמן שאין חיבור בפועל. ה-detection
מתבסס על בדיקת תהליך + חותמת זמן של קובץ הסטטוס, ולכן מתאושש גם
מקריסות קשות של ה-injector.
</p>

<h3>הסרה שמנקה אחרי עצמה</h3>
<p>
<code>uninstall.bat</code> בארבעה שלבים: עוצר את ה-tray וה-injector, מסיר
את ה-Startup entry ואת המפתח ב-<code>HKCU\...\Uninstall\ClaudeWordRTL</code>,
מנקה את משתנה הסביבה הישן של Auto-enable (אם הוא תואם לערך שלנו - <code>=9222</code> או <code>=0</code>), ומסיר את התלויות. המטרה: להסיר את כל מה שההתקנה יצרה. ערכים שהמשתמש הגדיר
ידנית באותו משתנה סביבה לא נמחקים.
</p>

<hr>

<h2 id="install">התקנה</h2>

<p>
<strong>דרישות:</strong> Windows 10/11, Microsoft Office desktop (לפחות אחת מבין Word, Excel, PowerPoint, או Outlook הקלאסי) עם תוסף Claude מותקן. תמיכת Outlook היא opt-in - מי שלא רוצה לחבר את Outlook יכול להתעלם מהפיצ'ר לחלוטין; הוא לא מופעל אוטומטית.
</p>

<blockquote>
<p><strong>צריך Node.js 16+. בלעדיו ההתקנה לא תמשיך וה-tray יישאר אדום.</strong></p>
<p>
אם לא מותקן (בדיקה: <code>node --version</code> ב-cmd), להוריד את גרסת ה-<strong>LTS</strong>
מ-<a href="https://nodejs.org/">nodejs.org</a> ולהתקין עם ה-defaults (Next-Next-Next).
לא דורש הרשאות admin. אחרי ההתקנה לפתוח cmd חדש ולהריץ <code>node --version</code> - צריך לראות <code>v16</code> ומעלה.
</p>
</blockquote>

<ol>
  <li><strong>להתקין Node.js 16+</strong> אם עוד לא מותקן (ראו למעלה).</li>
  <li>להוריד זיפ מ-<a href="https://github.com/asaf-aizone/Claude-for-Office-RTL-fix/releases">Releases</a>, או <code>git clone</code>.</li>
  <li>לחלץ לתיקייה שתישמר (למשל <code>C:\Tools\claude-office-rtl\</code>).</li>
  <li>לסגור את Word/Excel/PowerPoint אם הם פתוחים.</li>
  <li>דאבל-קליק על <strong><code>install.bat</code></strong>. ההתקנה רצה בארבעה שלבים בלי שאלות, ובסיומה אייקון ה-tray יעלה ליד השעון. הוא יעלה אוטומטית גם בלוגין הבא.</li>
</ol>

<p>
לוגים נכתבים ל-<code>install.log</code> ליד ה-installer. לא נדרשות הרשאות admin.
</p>

<p>
  <img src="docs/images/installer-done.png" alt="פלט install.bat בסיום מוצלח: 4 שלבים, ללא Auto-enable, ה-tray מתחיל אוטומטית" width="640">
</p>

<h3>עדכון לגרסה חדשה</h3>
<p>
קליק ימני על אייקון ה-tray &gt; <strong>Check for updates...</strong>. אם יש גרסה חדשה,
הדיאלוג יראה את מיקום ההתקנה הנוכחית שלך, יפתח את דף ההורדה בדפדפן,
ויפתח את תיקיית ההתקנה ב-Explorer. אחר כך:
</p>
<ol>
  <li>לחלץ את ה-zip מעל תיקיית ההתקנה (להחליף את כל הקבצים).</li>
  <li>לסגור את Word.</li>
  <li>להריץ את <code>install.bat</code> שוב. הוא יעצור את הטריי וה-injector הישנים,
      ידליק את החדשים עם הקוד המעודכן.</li>
</ol>

<h3>הסרה</h3>
<p>
דאבל-קליק על <strong><code>uninstall.bat</code></strong>.
</p>

<hr>

<h2>שימוש יום-יומי</h2>

<div dir="rtl" style="border:1px solid #d0d7de; background:#f6f8fa; padding:12px 16px; border-radius:6px;">
<strong>התזרים הנכון (Connect קודם, אז פותחים את הקובץ):</strong>
<ol>
  <li>אם Word / Excel / PowerPoint / Outlook פתוחים - שמרו וסגרו אותם.</li>
  <li>קליק ימני על אייקון ה-tray ליד השעון, ובוחרים <code>Connect Word</code> (או Excel / PowerPoint / Outlook).</li>
  <li>מאשרים בדיאלוג. ה-tray מפעיל את האפליקציה (ריקה) דרך ה-wrapper, והאייקון הופך לירוק.</li>
  <li><strong>רק עכשיו</strong> פותחים את הקובץ מתוך האפליקציה (File &gt; Open, Recent, או גרירה לחלון). הפאנל יעלה RTL.</li>
</ol>
<small><strong>למה הסדר חשוב:</strong> דגל ה-debug של WebView2 נקלט רק בעליית התהליך של Office. קובץ שנפתח <em>אחרי</em> שהאפליקציה הופעלה דרך ה-wrapper - יקבל RTL. קובץ שנפתח לפני - לא, וצריך לסגור ולהתחיל שוב.</small>
</div>

<details dir="rtl">
<summary>אם האפליקציה כבר פתוחה עם מסמך לא שמור (מסלול fallback)</summary>

<p>
ה-tray תומך גם בלחיצת Connect כשהאפליקציה כבר פתוחה - הוא ינסה לסגור אותה בעדינות, יפעיל מחדש דרך ה-wrapper, ויפתח מחדש את אותם מסמכים. המסלול הזה עובר דרך COM enumeration ויש בו יותר חלקים זזים, ולכן הוא <strong>לא מומלץ כברירת מחדל</strong> - הוא קיים בעיקר כדי לא להיתקע אם כבר התחלתם לעבוד:
</p>

<ol>
  <li>קליק ימני על האייקון. שורת הסטטוס בראש התפריט תראה <code>running without RTL</code> לאפליקציה שפתוחה. בוחרים את <strong>Connect Word</strong>, <strong>Connect Excel</strong>, <strong>Connect PowerPoint</strong> או <strong>Connect Outlook</strong> בהתאם.</li>
  <li>דיאלוג יקפוץ עם הסבר על מה שיקרה. אם יש מסמכים/חוברות עבודה/מצגות לא שמורים - תופיע אזהרה מפורשת. אישור סוגר את האפליקציה ומריץ אותה מחדש דרך ה-wrapper המתאים עם אותם קבצים.</li>
  <li>שורת הסטטוס של אותה אפליקציה הופכת ל-<code>connected</code> והאייקון ל<strong>ירוק</strong>.</li>
  <li>אם משהו נכשל (האפליקציה לא עלתה מחדש, או עלתה בלי המסמכים) - סגרו ידנית והשתמשו בתזרים המומלץ למעלה.</li>
</ol>
</details>

<p>
  <img src="docs/images/connect-dialog.png" alt="דיאלוג האישור של Connect - עם אזהרת UNSAVED" width="500">
</p>

<hr>

<h2>איך זה עובד?</h2>

<p>
לכל אפליקציית Office יש wrapper משלה (<code>word-wrapper.bat</code>,
<code>excel-wrapper.bat</code>, <code>powerpoint-wrapper.bat</code>). ה-wrapper מריץ את האפליקציה עם משתנה סביבה
<code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0</code>.
זה flag רשמי של Microsoft שפותח Chrome DevTools Protocol ב-localhost על
פורט דינמי. הערך <code>0</code> מאפשר ל-WebView2 לבחור פורט פנוי משלו לכל
תהליך, מה שנדרש כדי שכמה אפליקציות Office תוכלנה לרוץ בו-זמנית בלי
התנגשויות (בגרסת v0.1.x הפורט היה <code>9222</code> קבוע, מה שלא הספיק
לכמה תהליכים יחד). תהליך Node יחיד (<code>scripts/inject.js</code>) משתמש ב-<code>scripts/port-discovery.js</code> כדי לסרוק כל tick את כל תהליכי <code>msedgewebview2.exe</code> דרך <code>tasklist</code>, למפות אותם ל-LISTENING ports דרך <code>netstat</code>, ולבדוק כל פורט מועמד מול <code>/json/list</code> של CDP. עבור כל target הוא מזהה את האפליקציה דרך הפרמטר <code>_host_Info=</code> ב-URL של הפאנל, מתחבר דרך WebSocket, ומריץ <code>Runtime.evaluate</code>
כדי להזריק <code>&lt;style&gt;</code> ו-MutationObserver. לולאה של שתי שניות מזריקה מחדש
אם הפאנל טוען את עצמו.
</p>

<p>
ה-injector לא יוצר ולא שולח בקשות HTTP. הפאנל ממשיך לדבר ישירות מול
Anthropic כמו תמיד - ה-wrapper פשוט מוסיף flag ל-WebView2 ומחכה בצד.
</p>

<p>
ה-Connect לא תוקע את ה-UI: התפריט נסגר מיד, וה-state machine ממשיך ברקע
דרך timer. אם האפליקציה לא נסגרת תוך 10 שניות, מופיע דיאלוג OK/Cancel - OK
מחסל את התהליך ומריץ מחדש, Cancel משאיר את האפליקציה כפי שהיא.
</p>

<p>
מודל האיומים המלא: <a href="docs/security.md"><code>docs/security.md</code></a>.
</p>

<hr>

<h2>מה הכלי נוגע בו?</h2>

<table dir="rtl">
  <thead>
    <tr><th>משאב</th><th>גישה</th><th>למה</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>WebView2 של Word/Excel/PowerPoint</td>
      <td>קריאה דרך Chrome DevTools Protocol על localhost בפורט דינמי לכל אפליקציה</td>
      <td>לאתר את הפאנל של Claude בכל אחת משלוש האפליקציות ולהזריק CSS</td>
    </tr>
    <tr>
      <td>משתנה סביבה <code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code></td>
      <td>כתיבה רק ב-process scope של ה-wrapper הרלוונטי (Word/Excel/PowerPoint), שיורש לאפליקציה שהוא מפעיל. <strong>אין כתיבה ל-<code>HKCU\Environment</code></strong> מ-v0.1.4 ואילך</td>
      <td>לפתוח את ה-debug port ב-WebView2 ברגע הפעלה של אפליקציית Office. תהליכים אחרים של WebView2 לא רואים את המשתנה</td>
    </tr>
    <tr>
      <td><code>%TEMP%</code></td>
      <td>כתיבה של PID, סטטוס מצרפי (<code>claude-word-rtl.status</code>) וסטטוס לפי אפליקציה (<code>claude-office-rtl.apps.json</code>) של ה-injector</td>
      <td>למעקב אחר מצב ההזרקה מה-tray, להציג ארבע שורות סטטוס בתפריט, ולמנוע mass-kill של תהליכי Node</td>
    </tr>
    <tr>
      <td>Startup folder של המשתמש</td>
      <td>יצירת קיצור אחד לכניסה (<code>Claude for Word RTL Tray.lnk</code>; השם נשמר מ-v0.1.x לתאימות בעדכון)</td>
      <td>להפעיל את ה-tray אוטומטית בלוגין</td>
    </tr>
    <tr>
      <td><code>HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL</code></td>
      <td>כתיבה בהתקנה, מחיקה בהסרה. <code>DisplayName</code> נשאר "Claude for Word RTL Fix" לתאימות בעדכון מ-v0.1.x</td>
      <td>רישום הכלי ב-Windows Settings &gt; Apps כדי שיופיע ברשימת Installed apps ויהיה ניתן להסרה משם</td>
    </tr>
    <tr>
      <td>Word, Excel, PowerPoint (COM)</td>
      <td>קריאה בלבד, רק בזמן Connect, עבור האפליקציה שעליה לחצו</td>
      <td>למנות מסמכים/חוברות עבודה/מצגות פתוחים (<code>Documents</code>/<code>Workbooks</code>/<code>Presentations</code>) כדי לפתוח אותם מחדש אחרי relaunch</td>
    </tr>
    <tr>
      <td>Outlook (CDP בלבד, בלי COM)</td>
      <td>חיבור CDP רק אחרי Connect Outlook + אישור דיאלוג, בכפוף לדגל opt-in זמני. <strong>אין enumeration של תיבת הדואר דרך COM</strong> - מייל וקלנדר server-side וייפתחו לבד אחרי relaunch. ניתוק אוטומטי אחרי 15 דקות. ה-URL ב-לוג מצונזר (פרמטרים של tenant מוסרים)</td>
      <td>להחיל RTL על הפאנל של Claude בתוך Outlook הקלאסי. גישה זו מודעת לכך שתוכן מייל חשוף ב-DOM של הפאנל בזמן Summarize/Draft, ולכן מותנית בהסכמה מפורשת לכל הפעלה</td>
    </tr>
  </tbody>
</table>

<p>
<strong>בנוסף לטבלה למעלה, הכלי לא נוגע ב:</strong> file associations,
<code>Normal.dotm</code>, תבניות אחרות של Office, תוספים אחרים, או services.
ב-registry הוא נוגע רק במפתח אחד שבטבלה
(<code>HKCU\...\Uninstall\ClaudeWordRTL</code>). ב-v0.2.0 אין כתיבה
ל-<code>HKCU\Environment</code>; ההתקנה וההסרה רק <em>מנקות</em> שם את משתנה
ה-Auto-enable הישן אם הוא נשאר ממה-v0.1.x והערך תואם לאחד מהערכים שלנו.
</p>

<hr>

<h2>פרטיות</h2>

<ul>
  <li>בלי טלמטריה, אנליטיקס, או usage tracking.</li>
  <li>בלי חיבורי רשת יוצאים שהכלי יוזם.</li>
  <li>הכלי לא קורא את הפרומפטים שלך, את תשובות Claude, או את תוכן המסמכים. הקבצים היחידים שנכתבים לדיסק הם PID ו-status של ה-injector ב-<code>%TEMP%</code>, ולוגים אופציונליים של install ו-doctor.</li>
  <li>השיחה עם Claude עוברת ישירות בין WebView2 ל-Anthropic, בדיוק כמו בלי הכלי.</li>
</ul>

<h3>הערת אבטחה</h3>

<p>
כל עוד אפליקציית Office (Word/Excel/PowerPoint/Outlook) רצה דרך הכלי, ה-WebView2 שלה פותח debug port ב-localhost על פורט דינמי (אחד לכל תהליך WebView2 host של Office).
ה-port לא חשוף לרשת, אבל תהליך מקומי באותו user יכול להתחבר אליו (זה
המנגנון הסטנדרטי של Chrome DevTools Protocol; כל דפדפן מבוסס Chromium
שפותח debug port מתנהג אותו דבר). בפועל, כמעט כל מה שרץ בסשן שלכם יכול
כבר לקרוא את הזיכרון של אפליקציית Office. השימוש ב-<strong>Disconnect all</strong> מה-tray
כשסיימתם מנקה את ה-ports. על מחשבים משותפים או לא מהימנים - אל תריצו.
</p>

<p>
<strong>Outlook ספציפית (חדש ב-v0.3.0):</strong> בזמן שהפאנל של Claude מסכם
מייל או מנסח תגובה, תוכן המייל הופך לחלק מה-DOM של הפאנל ולכן חשוף לאותו
מנגנון CDP. החשיפה צרה בזמן (רק כשהפעולה רצה) אבל הקטגוריה רגישה יותר -
לכן Connect Outlook מציג דיאלוג opt-in, ה-injector מנתק אוטומטית אחרי 15 דקות,
ויש פריט תפריט ייעודי Disconnect Outlook only שמנתק רק את Outlook בלי
לפגוע ב-Word/Excel/PowerPoint. ראו את <a href="docs/security.md#outlook-specific-risks-and-mitigations">Outlook-specific risks and mitigations</a> ב-<code>docs/security.md</code> לפרטים מלאים.
</p>

<p>
דיווח על פגיעויות: <a href="SECURITY.md"><code>SECURITY.md</code></a>.
</p>

<hr>

<h2>שאלות ותשובות</h2>

<p><strong>האם Anthropic יחסמו אותי?</strong><br>
לא צפוי. הכלי רק משנה איך הפאנל נראה ב-DOM המקומי שלך. מה שאתה שולח
ל-Claude ומה שהוא מחזיר לא משתנה, ו-rate limits ו-guardrails לא
נוגעים בהם. עם זאת, כמו בכל כלי צד-שלישי - תנאי השימוש של Anthropic
הם הקובעים, והאחריות על השימוש היא שלך.</p>

<p><strong>מה הכלי עושה ל-Office?</strong><br>
לא משנה את Word/Excel/PowerPoint עצמם: בלי patch לתוסף, בלי שינוי טמפלטים, בלי
file associations. בהתקנה נוצרים קיצור בתיקיית Startup של המשתמש
ומפתח תחת <code>HKCU\...\Uninstall\ClaudeWordRTL</code> (כדי שהכלי יופיע
ב-Windows Settings &gt; Apps). שניהם מוסרים על ידי <code>uninstall.bat</code>,
וניתנים למחיקה ידנית גם כן.</p>

<p><strong>למה אני צריך Node.js?</strong><br>
ה-injector כתוב ב-Node כי הוא צריך לדבר עם CDP דרך WebSocket. בלי Node
האייקון יעלה אדום ויישאר אדום. אם אין לך Node מותקן, ההתקנה תסב תשומת
לבך לזה.</p>

<p><strong>עדכון של תוסף Claude ישבור את זה?</strong><br>
אפשר. ה-injector תלוי במבנה ה-DOM וב-URL pattern של הפאנל. אם Anthropic
ישנו משהו משמעותי, התיקון הוא בדרך כלל עדכון של סלקטור אחד. תפתחו issue
עם צילום מסך ונוציא patch.</p>

<p><strong>Office Online? Mac? Microsoft 365 תאגידי עם EDR?</strong><br>
Office Online (Word Online, Excel Online, PowerPoint Online) - לא, הכלי דורש WebView2 של Office desktop. Mac - לא, Windows
בלבד. תאגידי - לבדוק עם IT לפני הפעלה של debug port ב-Office. לא
מיועד ל-laptops תאגידיים סגורים. ראו את האזהרה למחשבים מנוהלי-ארגון בראש הקובץ.</p>

<p><strong>Word/Excel/PowerPoint פתוחים עם הרבה קבצים. Connect ייסגור אותם?</strong><br>
כן, אבל בעדינות - האפליקציה מתבקשת לשמור שינויים, הכלי מקבל את רשימת הקבצים הפתוחים דרך
COM (לפי האפליקציה: <code>Documents</code>/<code>Workbooks</code>/<code>Presentations</code>), וה-wrapper פותח את כולם מחדש. אם משהו לא נשמר, האפליקציה תשאל כרגיל.</p>

<p><strong>אין לי Git. אפשר בלי clone?</strong><br>
כן. להוריד זיפ מ-Releases, לחלץ, להריץ <code>install.bat</code>.</p>

<p><strong>איך בודקים אם יש גרסה חדשה?</strong><br>
הריצו <code>check-update.bat</code> או השתמשו בתפריט של הטריי: "Check for updates...".
הסקריפט משווה את הגרסה המקומית לגרסה האחרונה ב-GitHub דרך ה-API של release. ללא תלויות npm חיצוניות.</p>

<hr>

<h2>פתרון בעיות</h2>

<h3>אבחון מהיר - השתמשו ב-Claude Code, לא ב-Claude Chat</h3>

<p>
לפני שממשיכים לטבלת התסמינים למטה, שימו לב: לאבחון של הכלי הזה
<strong>השתמשו ב-Claude Code</strong> (<a href="https://claude.com/claude-code">claude.com/claude-code</a>),
לא ב-Claude Chat או בפאנל של Claude ב-Word. הסיבה: Claude Code רץ
מקומית על המחשב שלכם, קורא ישירות את <code>%TEMP%\claude-word-rtl.log</code>,
<code>doctor.log</code>, ואת <code>CLAUDE.md</code> של הפרויקט, ויכול להריץ
<code>netstat</code>, <code>curl</code>, ו-<code>tasklist</code> כחלק מהאבחון.
</p>

<p>
זרימת אבחון מומלצת: להתקין את Claude Code, לפתוח session בתיקיית
ההתקנה (<code>cd</code> לשם ואז <code>claude</code>), ולתאר את הבעיה
בעברית. Claude Code יקרא את הלוגים, יזהה את הגורם, ויציע תיקון.
Chat או הפאנל של Claude ב-Word לא רואים את הקבצים האלה, אז הם
ינחשו ויאכילו אתכם בצעדים כלליים שלא יעזרו במצב הספציפי.
</p>

<p>
לפני הכל - לפתוח את <strong>Show diagnostic log</strong> מהתפריט. הלוג ב-
<code>%TEMP%\claude-word-rtl.log</code> נחתך בכל הפעלה ומציג: targets שנמצאו
ב-CDP, אירועי attach, שגיאות <code>listTargets</code>. ב-90% מהמקרים הסיבה
ברורה משם.
</p>

<table dir="rtl">
  <thead>
    <tr><th>סימפטום</th><th>מה לעשות</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>האייקון לא עולה אחרי התקנה</td>
      <td>לבדוק שה-Startup entry נוצר: <code>Win+R</code>, <code>shell:startup</code>, לחפש את הקיצור</td>
    </tr>
    <tr>
      <td>האייקון נשאר אדום אחרי Connect</td>
      <td>פותחים <strong>Show diagnostic log</strong> מהתפריט. אם הלוג ריק או לא מובן - <code>doctor.bat</code> ולצרף ל-issue</td>
    </tr>
    <tr>
      <td>Connect לא סוגר את Word</td>
      <td>אחרי 10 שניות מופיע דיאלוג OK/Cancel. <strong>OK</strong> מחסל את התהליך ומריץ מחדש. <strong>Cancel</strong> משאיר את Word פתוח כדי לבדוק ידנית מה תוקע אותו</td>
    </tr>
    <tr>
      <td>תהליכי Node תקועים</td>
      <td><strong><code>cleanup.bat</code></strong> - מכוון לתהליכי Node שמריצים את <code>inject.js</code> של הכלי (זיהוי לפי command line), ולא לתהליכי Node אחרים שלא קשורים אליו</td>
    </tr>
    <tr>
      <td>האייקון אדום, Node לא מותקן</td>
      <td>לבדוק עם <code>node --version</code>. אם לא מותקן או מתחת ל-16, להוריד מ-<a href="https://nodejs.org/">nodejs.org</a></td>
    </tr>
    <tr>
      <td>הפאנל טוען והאייקון נשאר אדום</td>
      <td>ה-URL של הפאנל אולי השתנה. לפתוח issue עם צילום מסך של ה-URL ב-DevTools, ולצרף את <code>%TEMP%\claude-word-rtl.log</code></td>
    </tr>
    <tr>
      <td>RTL לא מופיע אחרי Connect</td>
      <td>הריצו <code>doctor.bat</code>. הוא מבצע 19 בדיקות (החל מ-v0.3.0) הכוללות סריקת פורטי CDP דינמיים פעילים של Office WebView2 דרך <code>tasklist</code> + <code>netstat</code>, וזיהוי targets של Claude לפי אפליקציה. ארבע הבדיקות האחרונות (16-19) ייעודיות ל-Outlook: התקנה, תהליך רץ, target CDP, ומצב ב-<code>apps.json</code>. כולן <code>INFO</code> כי Outlook הוא opt-in. בנוסף לוגים של ה-injector ב-<code>%TEMP%\claude-word-rtl.log</code> (נחתך בכל הפעלה) מציגים אילו ports נסרקו ואילו targets אותרו. אם <code>doctor.bat</code> מראה רשימת פורטים ריקה - האפליקציה לא נפתחה דרך ה-wrapper שלה (פתיחה ישירה מאייקון Word/Excel/PowerPoint/Outlook לא מפעילה את ה-debug port). השתמשו ב-Connect המתאים מהטריי.</td>
    </tr>
  </tbody>
</table>

<hr>

<h2>מגבלות ידועות</h2>

<ul>
  <li>debug port של WebView2 ב-localhost (פורט דינמי) לא מאומת - כל תהליך מקומי באותו user יכול להתחבר. ראו "הערת אבטחה".</li>
  <li>Microsoft 365 תאגידי עם EDR/DLP יכול לחסום את דגל ה-WebView2. הכלי לא מיועד ללפטופים ארגוניים סגורים.</li>
  <li>עדכון של תוסף Claude שמחליף את ה-DOM יכול לשבור את ההזרקה עד patch. יישלח release מתוקן.</li>
  <li><strong>Mac (macOS) לא נתמך ולא יהיה נתמך.</strong> Office ל-Mac משתמש ב-WKWebView במקום WebView2, ושכבת ה-launcher כולה (bat, vbs, ps1) היא Windows-only. Office Online (Word/Excel/PowerPoint Online) גם לא נתמך.</li>
  <li>גרסאות של תוסף Claude שלא משתמשות ב-WebView2 (למשל Electron עצמאי) - לא נתמכות.</li>
</ul>

<hr>

<h2>תרומה לפרויקט</h2>

<p>
Issues ו-PRs מתקבלים בברכה. לבאגים של תצוגה - selector שדלף, יישור שבור -
תפתחו issue עם:
</p>

<ul>
  <li>שחזור קצר: מה עשית, מה ציפית, מה קרה.</li>
  <li>צילום מסך של הפאנל.</li>
  <li>ה-<code>doctor.log</code> שלך.</li>
</ul>

<h2>קרדיטים</h2>

<ul>
  <li>נוצר על ידי <strong>Asaf Abramzon</strong> - <a href="https://www.linkedin.com/in/asaf-abramzon-7a2b61180/">LinkedIn</a> · <a href="https://github.com/asaf-aizone">GitHub</a>.</li>
  <li><a href="https://github.com/cyrus-and/chrome-remote-interface"><code>chrome-remote-interface</code></a> - CDP client.</li>
</ul>

<h2>כתב ויתור</h2>

<p>
כלי open-source עצמאי. לא מסונף ל-Anthropic או Microsoft, לא מאושר על ידן,
ולא מכיל קוד שלהן. "Claude" סימן מסחרי של Anthropic, PBC. "Microsoft",
"Word", "Excel" ו-"PowerPoint" סימנים מסחריים של Microsoft Corporation.
</p>

<h2>מסמכים נוספים</h2>

<ul>
  <li><a href="CHANGELOG.md"><code>CHANGELOG.md</code></a> - יומן שינויים לכל גרסה.</li>
  <li><a href="README.he.md"><code>README.he.md</code></a> - גרסה עברית תמציתית (טקסט בלבד, ללא תמונות).</li>
  <li><a href="SECURITY.md"><code>SECURITY.md</code></a> - מדיניות דיווח פגיעויות.</li>
  <li><a href="docs/security.md"><code>docs/security.md</code></a> - מודל האיומים המלא.</li>
</ul>

<h2>רישיון</h2>

<p>Apache License 2.0, ראו <a href="LICENSE"><code>LICENSE</code></a>.</p>

</div>

<hr>

<details>
<summary><strong>English version</strong></summary>

<h1>Claude for Office RTL Fix (Word, Excel, PowerPoint, Outlook)</h1>

<p>
RTL fix for the Claude panel in Microsoft Word, Excel, PowerPoint, and Outlook. Local-only, no telemetry, Apache 2.0.
</p>

<p>
<strong>New in v0.3.0:</strong> classic Outlook support behind a hardened
opt-in model (explicit per-launch consent dialog, 15-minute
auto-disconnect, URL redaction in the diagnostic log). The other three
apps are unchanged. See the <a href="#outlook-section-en">Outlook (opt-in)</a>
section below.
</p>

<p>
<strong>Not an official Anthropic or Microsoft add-in.</strong> Independent open-source tool. The GitHub repository was renamed in v0.2.1 to <code>Claude-for-Office-RTL-fix</code> (was <code>Claude-for-word-RTL-fix</code>). GitHub keeps a permanent redirect from the old name, so existing clones, bookmarks, and v0.1.x install URLs continue to work.
</p>

<blockquote>
<p>
<strong>Windows only.</strong> This tool does not work on macOS or Linux. The Claude
add-in for Office is built on Microsoft's WebView2 runtime, which is Windows-only. Office
for Mac uses WKWebView, which does not expose the same debugging interface, and the
entire launcher stack (batch, VBS, PowerShell, registry, Startup folder) is
Windows-specific. If you are on Mac, this tool has no port.
</p>
</blockquote>

<h2>What it does</h2>

<p>
Anthropic's official Claude add-in for Office (Word, Excel, PowerPoint) doesn't render Hebrew right-to-left.
Bullets land on the wrong side, alignment is reversed, punctuation ends up in
weird places. This tool attaches to each Office app's WebView2 panel via Chrome
DevTools Protocol, injects a small stylesheet plus a MutationObserver, and
flips the panel to correct RTL. The Claude model, the Anthropic API, and the Office apps
themselves are untouched - only the panel's local DOM, and only while the panel
is open.
</p>

<p>
The reason this tool exists is accessibility: Hebrew speakers need RTL rendering
to read Claude's responses inside the add-in. Without it the panel shows Hebrew
LTR with broken bidi handling, which makes it effectively unusable. The fix is
an accessibility adaptation applied to locally-rendered output, in the same
spirit as a user stylesheet (Stylus/Stylish), a screen reader, or a dark-mode
injector. The underlying Service is not changed in any way.
</p>

<h2>Features</h2>

<ul>
  <li><strong>Instant RTL across Word, Excel, and PowerPoint</strong> - direction, alignment, bullets, tables. <code>&lt;pre&gt;</code> and <code>&lt;code&gt;</code> stay LTR so source code isn't corrupted. One injector serves all three apps; you can have all of them open simultaneously with the panel RTL in each.</li>
  <li><strong>Tray-icon control</strong> - a small rounded-square icon near the clock: a white <strong>O</strong> (Office) and an RTL arrow on a status-colored background. Gray = starting. Red = not attached. Green = at least one Office app is connected with RTL. Right-click for three per-app status labels (Word/Excel/PowerPoint), three Connect items (Connect Word / Connect Excel / Connect PowerPoint), Disconnect all, Show diagnostic log, Check for updates, Uninstall, Exit. No Start Menu entry pretending to be Word. See the <a href="docs/images/tray-icon-red.png">red (disconnected)</a> and <a href="docs/images/tray-icon-green.png">green (connected)</a> states.</li>
  <li><strong>Per-process activation</strong> - Connect Word/Excel/PowerPoint each launches the matching Office app through its wrapper, with the WebView2 debug flag set in the wrapper's process scope only. The flag is inherited by Office but never seen by Teams, Outlook, Edge, or any other WebView2 host on your account.</li>
  <li><strong>Non-blocking Connect</strong> - the menu closes immediately, work proceeds on a background timer state machine. If the Office app doesn't close within 10 seconds, an OK/Cancel dialog offers force-kill or abort.</li>
  <li><strong>Minimal footprint</strong> - no file associations, no <code>Normal.dotm</code> changes, no Office templates, no services. Install creates two per-user items: a Startup-folder shortcut (so the tray auto-launches at login; filename <code>Claude for Word RTL Tray.lnk</code> retained for v0.1.x upgrade compat) and an <code>HKCU\...\Uninstall\ClaudeWordRTL</code> registry key (so the tool appears in Windows Settings &gt; Apps). Both are removed by <code>uninstall.bat</code>.</li>
  <li><strong>Hebrew typography cleanup</strong> - em-dash and en-dash become hyphen. Arrow glyphs become commas. Input fields and code blocks are left alone.</li>
  <li><strong>Crash-safe status</strong> - if the injector dies without cleanup, the tray detects stale state (process + status-file timestamp) and flips to red, to reduce cases where the icon shows green without an actual connection.</li>
  <li><strong>Diagnostic log</strong> - <code>%TEMP%\claude-word-rtl.log</code>, accessible from the tray menu. Truncated on each injector start. Shows discovered Office CDP ports, per-target app identification, attach events, errors.</li>
  <li><strong>Clean uninstall</strong> - 4-step <code>uninstall.bat</code>: stop tray/injector, remove Startup entry + <code>HKCU\...\Uninstall\ClaudeWordRTL</code>, clear any v0.1.x legacy Auto-enable env var (only if it still matches one of our known values, <code>=9222</code> or <code>=0</code>; user-modified values are preserved), prune deps. Aims to remove everything the installer created.</li>
</ul>

<h3>Tray menu</h3>

<table>
  <thead>
    <tr><th>Item</th><th>When available</th><th>What it does</th></tr>
  </thead>
  <tbody>
    <tr><td>Word: ... / Excel: ... / PowerPoint: ... / Outlook: ...</td><td>Always (four disabled labels at the top)</td><td>Read-only per-app status labels. Each can read <code>connected</code>, <code>not running</code>, <code>running without RTL</code>, or <code>error</code>. Refreshed every 2s from <code>%TEMP%\claude-office-rtl.apps.json</code>.</td></tr>
    <tr><td>Connect Word</td><td>Anytime</td><td>Relaunches Word through <code>word-wrapper.bat</code> with the debug-port enabled (dynamic port). If Word is already open, enumerates documents via <code>Word.Application</code> COM, asks for confirmation, and reopens them.</td></tr>
    <tr><td>Connect Excel</td><td>Anytime</td><td>Same flow for Excel via <code>excel-wrapper.bat</code>; enumerates <code>Workbooks</code>.</td></tr>
    <tr><td>Connect PowerPoint</td><td>Anytime</td><td>Same flow for PowerPoint via <code>powerpoint-wrapper.bat</code>; enumerates <code>Presentations</code>.</td></tr>
    <tr><td>Connect Outlook</td><td>Always (opt-in item)</td><td>Shows a dedicated content-exposure warning dialog whose default-focused button is Cancel - a stray Enter does not opt in. On OK, invokes <code>outlook-wrapper.bat</code>, which writes a per-launch opt-in flag (<code>%TEMP%\claude-office-rtl.outlook-optin</code>) and launches classic Outlook. The injector blocks Outlook by default; the flag is the required gate. Refuses to run if New Outlook (<code>olk.exe</code>) is detected. No COM enumeration (mail and calendar are server-side and reappear after relaunch).</td></tr>
    <tr><td>Disconnect Outlook only</td><td>Only when Outlook is <code>connected</code></td><td>Drops the Outlook CDP attachment without closing Outlook itself and without affecting Word/Excel/PowerPoint. IPC: writes a request file (<code>claude-office-rtl.disconnect-outlook.request</code>) that the injector polls each tick; the injector closes the Outlook CDP client(s), clears the per-Outlook auto-disconnect timer, and revokes the opt-in flag. Use this between mail sessions instead of relying on the 15-minute timer or Disconnect all.</td></tr>
    <tr><td>Disconnect all</td><td>Anytime</td><td>Universal recovery button. Stops any in-flight Connect timers, closes every open Word/Excel/PowerPoint (graceful + force fallback), kills the injector, cleans up state files, revokes the Outlook opt-in flag. <strong>Does NOT close Outlook itself</strong> (it is opt-in and the user may have opened it just to read mail with no intent to use RTL); killing the injector already drops the CDP attachment. If anything went wrong, clicking this always returns to a clean slate.</td></tr>
    <tr><td>Show diagnostic log</td><td>Anytime</td><td>Opens <code>%TEMP%\claude-word-rtl.log</code> in the default editor.</td></tr>
    <tr><td>Check for updates...</td><td>Anytime</td><td>Runs <code>check-update.js</code> and shows the result in a dialog. When a newer release exists, a one-click button opens the download page in the default browser.</td></tr>
    <tr><td>Uninstall...</td><td>Anytime</td><td>Runs <code>uninstall.bat</code> after confirmation.</td></tr>
    <tr><td>Exit</td><td>Anytime</td><td>Stops the tray icon.</td></tr>
  </tbody>
</table>

<h3>Activation via Connect (no persistent env var)</h3>

<p>
v0.1.4 removed the persistent Auto-enable toggle and v0.2.0 keeps it
removed. The reason: Auto-enable wrote
<code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code> at <code>HKCU\Environment</code> level,
which is read by <em>every</em> WebView2 host running under your user (Teams,
the new Outlook, Edge WebView, the OneDrive UI). Enterprise EDR products
(Microsoft Defender for Endpoint, CrowdStrike Falcon, SentinelOne,
Sophos) treat unexpected modifications of WebView2 browser arguments as
a credential-theft signal, and a v0.1.x field incident triggered host
isolation on a managed device.
</p>

<p>
v0.2.0 keeps activation Connect-only. Each Connect Word/Excel/PowerPoint
click invokes the matching wrapper
(<code>word-wrapper.bat</code>, <code>excel-wrapper.bat</code>,
<code>powerpoint-wrapper.bat</code>), which sets the WebView2 debug flag in
its own process scope. The flag is inherited only by the Office app the
wrapper launches, and is invisible to any other WebView2 host on the
account. The functional behavior of the old Auto-enable toggle (RTL
ready every time you open Office) is preserved without writing to the
registry.
</p>

<p>
On uninstall, if a v0.1.x Auto-enable env var is still set, it's cleared
only if its value matches one of our known strings (<code>=9222</code>
or <code>=0</code>); any user-modified value is preserved.
</p>

<h3 id="outlook-section-en">Outlook (opt-in, new in v0.3.0)</h3>

<p>
v0.3.0 adds classic Outlook to the supported set, but under a stricter
security model than the other three apps. Two reasons:
</p>

<ol>
  <li>When the user asks Claude to "Summarize this email" or "Draft a reply", the active email's content becomes part of the panel DOM. The same CDP attach that lets the injector apply RTL CSS also lets any local process under the same user account read that DOM. While the operation is in flight, the mail content is exposed to whatever else is running in your session.</li>
  <li>Mail can include credentials, MFA codes, legal documents, tenant identifiers. Qualitatively more sensitive than a Word document panel showing content the user explicitly pasted.</li>
</ol>

<p>
Outlook therefore has four protections that do <strong>not</strong> apply to Word/Excel/PowerPoint:
</p>

<ul>
  <li><strong>Explicit per-launch opt-in.</strong> The injector permanently blocks Outlook unless a per-launch flag file (<code>%TEMP%\claude-office-rtl.outlook-optin</code>) is present. The flag is written only by <code>outlook-wrapper.bat</code> after the user explicitly clicks Connect Outlook and accepts a warning dialog whose default-focused button is Cancel. The flag is never persisted across injector restarts - the injector clears it at startup, so a previous session's consent never carries silently into a new session.</li>
  <li><strong>No auto-launch.</strong> The tray auto-relaunches the injector when an Office app is up but the injector has died. Outlook is excluded from this trigger: if Outlook is the only Office app running, the injector stays down. A user who opens Outlook just to read mail does not accidentally bring CDP attach back online.</li>
  <li><strong>15-minute auto-disconnect.</strong> Even if the user forgets to disconnect, the injector closes the Outlook CDP client after 15 minutes of continuous attachment and revokes the opt-in flag. The user must click Connect Outlook again to start a new session. Tunable in source via <code>OUTLOOK_AUTO_DISCONNECT_MIN</code> in <code>scripts/inject.js</code>.</li>
  <li><strong>URL redaction in the diagnostic log.</strong> The Office add-in URL contains an <code>et=</code> parameter holding base64-encoded tenant metadata (account id, tenant id, expiry). For Outlook only, the injector strips every query parameter except <code>_host_Info=</code> from logged URLs, so <code>%TEMP%</code> no longer leaks tenant identifiers. Word/Excel/PowerPoint URLs are logged verbatim, unchanged from v0.2.x.</li>
</ul>

<p>
Plus a dedicated <strong>Disconnect Outlook only</strong> tray item that drops the Outlook CDP attachment without affecting Word/Excel/PowerPoint sessions still in flight.
</p>

<p>
<strong>New Outlook (<code>olk.exe</code>) is intentionally not supported.</strong> The M0 probe deferred it, and both <code>outlook-wrapper.bat</code> and Connect Outlook refuse to launch if New Outlook is detected, to avoid colliding on shared per-user state.
</p>

<p>
Full design rationale: <a href="docs/security.md#outlook-specific-risks-and-mitigations">Outlook-specific risks and mitigations</a> in <code>docs/security.md</code>, and the design plan in <a href="docs/OUTLOOK-EXPANSION-PLAN.md"><code>docs/OUTLOOK-EXPANSION-PLAN.md</code></a>.
</p>

<h2 id="install-en">Install</h2>

<p>
<strong>Requirements:</strong> Windows 10/11, Microsoft Office desktop (at least one of Word, Excel, PowerPoint, or classic Outlook) with Claude
add-in installed. Outlook support is opt-in - users who do not want it can ignore the feature entirely; it is never auto-activated.
</p>

<blockquote>
<p><strong>Node.js 16+ is required. Without it the installer stops and the tray stays red.</strong></p>
<p>
If not installed (check with <code>node --version</code> in cmd), download the
<strong>LTS</strong> installer from <a href="https://nodejs.org/">nodejs.org</a> and run it
with the defaults (Next-Next-Next). No admin rights needed. Open a new cmd afterward
and run <code>node --version</code> - you should see <code>v16</code> or higher.
</p>
</blockquote>

<ol>
  <li><strong>Install Node.js 16+</strong> if you don't have it already (see above).</li>
  <li>Download the zip from <a href="https://github.com/asaf-aizone/Claude-for-Office-RTL-fix/releases">Releases</a> or <code>git clone</code>.</li>
  <li>Extract to a folder you'll keep (e.g. <code>C:\Tools\claude-office-rtl\</code>).</li>
  <li>Close Word, Excel, and PowerPoint if any of them are open.</li>
  <li>Double-click <strong><code>install.bat</code></strong>. The installer runs 4 steps with no prompts; the tray icon appears near the clock when it finishes, and will launch automatically on every login.</li>
</ol>

<p>Logs go to <code>install.log</code> next to the installer. No admin rights needed.</p>

<h3>Updating to a newer version</h3>
<p>
Right-click the tray icon &gt; <strong>Check for updates...</strong>. If a newer
version exists, the dialog shows your current install folder, opens the
download page in the browser, and opens the install folder in Explorer.
Then:
</p>
<ol>
  <li>Extract the zip over your install folder (overwriting all files).</li>
  <li>Close Word, Excel, and PowerPoint.</li>
  <li>Run <code>install.bat</code> again. It stops the old tray and
      injector and starts the new ones with the updated code.</li>
</ol>

<h3>Uninstall</h3>
<p>Double-click <strong><code>uninstall.bat</code></strong>.</p>

<h2>Daily use</h2>

<blockquote>
<strong>Recommended flow (Connect first, then open the file):</strong>
<ol>
  <li>If Word / Excel / PowerPoint / Outlook are already open, save and close them.</li>
  <li>Right-click the tray icon near the clock and pick <code>Connect Word</code> (or Excel / PowerPoint / Outlook).</li>
  <li>Approve the dialog. The tray launches the empty app through the wrapper and the icon turns green.</li>
  <li><strong>Only now</strong> open your file from inside the app (File &gt; Open, Recent, or drag into the window). The panel comes up RTL.</li>
</ol>
<small><strong>Why the order matters:</strong> the WebView2 debug flag is inherited at process start, so a file opened <em>after</em> the app is launched through the wrapper gets RTL; a file opened before does not, and the app has to be closed and restarted.</small>
</blockquote>

<details>
<summary>If the app is already open with unsaved work (fallback path)</summary>

<p>
The tray also supports clicking Connect while the app is already open - it will try to close it gracefully, relaunch through the wrapper, and reopen the same documents. This path goes through COM enumeration and has more moving parts, so it is <strong>not the recommended default</strong>. It exists so you do not get stuck if you have already started working:
</p>

<ol>
  <li>Right-click the tray icon. The status label at the top of the menu for that app will read <code>running without RTL</code>. Pick <strong>Connect Word</strong>, <strong>Connect Excel</strong>, <strong>Connect PowerPoint</strong>, or <strong>Connect Outlook</strong> as appropriate.</li>
  <li>The tool enumerates your open documents/workbooks/presentations, asks for confirmation, reopens them through the matching wrapper.</li>
  <li>The status label flips to <code>connected</code> and the icon goes <strong>green</strong>.</li>
  <li>If anything fails (the app does not come back, or comes back without the documents) - close manually and use the recommended flow above.</li>
</ol>
</details>

<h2>How it works</h2>

<p>
Each Office app has its own wrapper (<code>word-wrapper.bat</code>,
<code>excel-wrapper.bat</code>, <code>powerpoint-wrapper.bat</code>). The wrapper launches the
Office app with
<code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0</code>,
a Microsoft-documented WebView2 flag. The value <code>0</code> means
"WebView2 picks a free dynamic port per process", so multiple Office apps
launched through their respective wrappers each get their own debug
surface without colliding (v0.1.x used a fixed <code>9222</code>, which
prevented multi-app support). A single Node process (<code>scripts/inject.js</code>)
uses <code>scripts/port-discovery.js</code> on every 2-second tick to enumerate
the active CDP surfaces: walk <code>tasklist</code> for
<code>msedgewebview2.exe</code> PIDs, map each PID to its LISTENING port via
<code>netstat</code>, and probe every candidate's <code>/json/list</code>. For each
Claude target it identifies the app from the <code>_host_Info=</code> URL
parameter (Word, Excel, or Powerpoint), opens a WebSocket to that
target's CDP endpoint, and calls <code>Runtime.evaluate</code> to inject a
<code>&lt;style&gt;</code> element and a MutationObserver.
</p>

<p>
All activity is local. The only outbound traffic is normal Claude-to-Anthropic
traffic, unchanged, routed by Office itself.
</p>

<p>Full threat model: <a href="docs/security.md"><code>docs/security.md</code></a>.</p>

<h2>What the tool accesses</h2>

<table>
  <thead>
    <tr><th>Resource</th><th>Access</th><th>Why</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>WebView2 of Word/Excel/PowerPoint</td>
      <td>Read via Chrome DevTools Protocol on a localhost dynamic port (one per Office WebView2 host)</td>
      <td>Locate the Claude panel in each app and inject CSS</td>
    </tr>
    <tr>
      <td><code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code> env var</td>
      <td>Wrapper process scope only. Each wrapper sets the variable just for itself; the launched Office app inherits it. <strong>Never written to <code>HKCU\Environment</code></strong> in v0.1.4 or v0.2.0</td>
      <td>Open the WebView2 debug port at Office-app launch. Other WebView2 hosts on the account never see the flag</td>
    </tr>
    <tr>
      <td><code>%TEMP%</code></td>
      <td>Writes injector PID, aggregate status (<code>claude-word-rtl.status</code>) and per-app status (<code>claude-office-rtl.apps.json</code>)</td>
      <td>Track injection state from the tray, render per-app status labels, prevent mass-kill of Node processes</td>
    </tr>
    <tr>
      <td>User Startup folder</td>
      <td>Creates one shortcut (<code>Claude for Word RTL Tray.lnk</code>; filename retained for v0.1.x upgrade compat)</td>
      <td>Launch the tray at login</td>
    </tr>
    <tr>
      <td><code>HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL</code></td>
      <td>Written on install, removed on uninstall. <code>DisplayName</code> still "Claude for Word RTL Fix" for v0.1.x upgrade compat</td>
      <td>Register the tool in Windows Settings &gt; Apps &gt; Installed apps so it can be uninstalled from there</td>
    </tr>
    <tr>
      <td>Word, Excel, PowerPoint (COM)</td>
      <td>Read-only, only during Connect, only for the app you clicked</td>
      <td>Enumerate open documents/workbooks/presentations (<code>Documents</code> / <code>Workbooks</code> / <code>Presentations</code>) to reopen them after relaunch</td>
    </tr>
    <tr>
      <td>Outlook (CDP only, no COM)</td>
      <td>CDP attach only after Connect Outlook + dialog confirmation, gated on a per-launch opt-in flag. <strong>No COM mailbox enumeration</strong> - mail and calendar are server-side and reappear on relaunch. Auto-disconnects after 15 minutes. URL parameters other than <code>_host_Info=</code> are redacted in the log</td>
      <td>Apply RTL to the Claude panel inside classic Outlook. This row exists because mail content enters the panel DOM during Summarize/Draft and is therefore exposed via the same CDP attach; the per-launch consent model is the gate</td>
    </tr>
  </tbody>
</table>

<p>
<strong>Beyond the table above, the tool does not touch:</strong> file
associations, <code>Normal.dotm</code>, any Word/Excel/PowerPoint template,
other Office add-ins, or Windows services. The only registry key it
writes is the one in the table (<code>HKCU\...\Uninstall\ClaudeWordRTL</code>).
v0.2.0 does not write <code>HKCU\Environment</code>; install and uninstall
only <em>clear</em> the legacy Auto-enable env var there if it remains
from v0.1.x and matches one of our known values.
</p>

<h2>Privacy</h2>

<ul>
  <li>No telemetry, analytics, or usage tracking.</li>
  <li>No outbound network connections initiated by this tool.</li>
  <li>No collection, storage, or logging of your prompts, Claude's responses, or document content. The only files written to disk are the injector's PID and status in <code>%TEMP%</code>, plus optional install/doctor logs.</li>
  <li>No third-party services. Conversations with Claude go directly between Word's WebView2 and Anthropic, exactly as without this tool.</li>
</ul>

<h3>Security note</h3>
<p>
While an Office app (Word, Excel, PowerPoint, or Outlook) runs via this tool,
its WebView2 host opens a debug port on a dynamic localhost port
(one per Office WebView2 host process). Any local process on the same user
can connect and read the panel's DOM. The port is localhost-only, but unauthenticated.
Use <strong>Disconnect all</strong> when done. Don't run on shared or untrusted machines.
</p>
<p>
<strong>Outlook specifically (new in v0.3.0):</strong> while Claude is
summarizing an email or drafting a reply, the mail content enters the
panel DOM and is exposed via the same CDP mechanism. The exposure
window is narrow (only while the operation is running) but the content
class is more sensitive than for document panels - hence Connect
Outlook gates on a per-launch consent dialog, the injector
auto-disconnects after 15 minutes, and a dedicated Disconnect Outlook
only menu item drops just the Outlook attachment without affecting
Word/Excel/PowerPoint. See
<a href="docs/security.md#outlook-specific-risks-and-mitigations">Outlook-specific risks and mitigations</a>
in <code>docs/security.md</code> for the full design.
</p>

<h2>FAQ</h2>

<p><strong>Will Anthropic block me?</strong><br>
Not expected. The tool only changes how the panel renders in your local DOM.
It doesn't alter what you send, what Claude returns, rate limits, or
guardrails. Your use of Claude remains governed by Anthropic's Terms of Service
and Usage Policy regardless of this tool. The tool does not change what you
send to Claude or what Claude sends back; it only restyles Claude's
already-rendered output in your local browser. If Anthropic's terms ever change
to restrict client-side modifications, comply with their terms over this tool.</p>

<p><strong>Does it modify Office?</strong><br>
It doesn't modify Word, Excel, or PowerPoint themselves: no template changes, no add-in patches, no
file-association changes. Install does create two per-user items: a
Startup-folder shortcut and an <code>HKCU\...\Uninstall\ClaudeWordRTL</code>
registry key (so the tool appears in Windows Settings &gt; Apps). Both are
removed by <code>uninstall.bat</code> and can also be cleaned up by hand.</p>

<p><strong>Why do I need Node.js?</strong><br>
The injector is written in Node because it needs to talk to CDP over
WebSocket. Without Node, the tray stays red.</p>

<p><strong>Will a Claude add-in update break it?</strong><br>
Possibly. The injector depends on the panel's DOM structure and URL pattern.
If Anthropic ships a significant change, the fix is usually a one-selector
update. Open an issue with a screenshot.</p>

<p><strong>Office Online, Mac, corporate M365 with EDR?</strong><br>
Office Online (Word/Excel/PowerPoint Online) - no, requires Office desktop's WebView2. Mac - no, Windows only.
Corporate - check with your IT team before enabling a WebView2 debug port on Office.
Not intended for sealed corporate laptops; see the EDR warning at the top of this README.</p>

<p><strong>Word, Excel, or PowerPoint is open with many files. Will Connect close them?</strong><br>
Yes, gracefully. The Office app is asked to save changes, the open-files list is captured
via the matching COM collection (<code>Documents</code>/<code>Workbooks</code>/<code>Presentations</code>),
and the wrapper reopens all of them. If something wasn't saved, the app prompts as usual.</p>

<p><strong>How do I check for a newer version?</strong><br>
Run <code>check-update.bat</code> or use the tray menu's "Check for updates..." item.
The script compares the local version to the latest GitHub release via the API. No npm dependencies.</p>

<h2>Troubleshooting</h2>

<h3>Quick diagnosis - use Claude Code, not Claude Chat</h3>

<p>
Before walking the symptom table below, note: for troubleshooting this
tool, <strong>use Claude Code</strong>
(<a href="https://claude.com/claude-code">claude.com/claude-code</a>), not
Claude Chat or Claude's in-Word panel. Reason: Claude Code runs
locally on your machine, reads
<code>%TEMP%\claude-word-rtl.log</code>, <code>doctor.log</code>, and
the project's <code>CLAUDE.md</code> directly, and can execute
<code>netstat</code>, <code>curl</code>, and <code>tasklist</code>
as part of its diagnosis.
</p>

<p>
Recommended flow: install Claude Code, open a session in the install
folder (<code>cd</code> into it, then <code>claude</code>), and describe
the problem. Claude Code reads the logs, identifies the cause, and
proposes a fix. Chat and the in-Word panel cannot see those files, so
they will guess from general knowledge and feed you generic steps that
often do not match the specific failure mode.
</p>

<p>
Always start with <strong>Show diagnostic log</strong> from the tray menu.
The log at <code>%TEMP%\claude-word-rtl.log</code> truncates on each injector
launch and shows discovered CDP targets, attach events, and <code>listTargets</code>
errors. 90% of issues are obvious from there.
</p>

<table>
  <thead>
    <tr><th>Symptom</th><th>What to do</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>Icon doesn't appear after install</td>
      <td>Check the Startup entry was created: <code>Win+R</code>, <code>shell:startup</code></td>
    </tr>
    <tr>
      <td>Icon stays red after Connect</td>
      <td>Open <strong>Show diagnostic log</strong>. If unclear, run <code>doctor.bat</code> and attach to an issue</td>
    </tr>
    <tr>
      <td>Connect doesn't close Word</td>
      <td>After 10s a dialog appears: <strong>OK</strong> force-kills and relaunches, <strong>Cancel</strong> leaves Word alone</td>
    </tr>
    <tr>
      <td>Lingering Node processes</td>
      <td><strong><code>cleanup.bat</code></strong> - targets Node processes running this tool's <code>inject.js</code> (matched by command line). Unrelated Node processes are left alone.</td>
    </tr>
    <tr>
      <td>Icon red, Node not installed</td>
      <td>Check with <code>node --version</code>. If missing or below 16, install from <a href="https://nodejs.org/">nodejs.org</a></td>
    </tr>
    <tr>
      <td>Panel loads but icon stays red</td>
      <td>The panel URL may have changed. Open an issue with a screenshot of the URL from DevTools and the diagnostic log</td>
    </tr>
    <tr>
      <td>RTL does not appear after Connect</td>
      <td>Run <code>doctor.bat</code>. As of v0.3.0 it performs 19 checks including a sweep of active Office WebView2 ports (via <code>tasklist</code> + <code>netstat</code>) and per-app Claude target enumeration. The four new checks (16-19) are Outlook-specific: install path, process running, CDP target, and apps.json state. All four are <code>INFO</code> because Outlook is opt-in. The injector log at <code>%TEMP%\claude-word-rtl.log</code> (truncated on each launch) shows which ports were probed and which targets matched. If <code>doctor.bat</code> reports an empty port list, the Office app was launched directly rather than through its wrapper, which does not enable the debug port; use the matching Connect item from the tray instead.</td>
    </tr>
  </tbody>
</table>

<h2>Known limitations</h2>

<ul>
  <li>The WebView2 debug port (a dynamic localhost port per Office WebView2 host) is unauthenticated. See Security note.</li>
  <li>Corporate M365 with EDR/DLP may block the WebView2 flag. Not intended for sealed corporate laptops.</li>
  <li>A Claude add-in update that changes the panel DOM can break injection until a patch is released.</li>
  <li><strong>Mac (macOS) is not supported and will not be supported.</strong> Office for Mac uses WKWebView instead of WebView2, and the entire launcher stack (bat, vbs, ps1) is Windows-only. Office Online (Word/Excel/PowerPoint Online) is also not supported.</li>
  <li>Custom builds of the Claude add-in (standalone Electron, alternate WebView) aren't supported.</li>
</ul>

<h2>Contributing</h2>

<p>
Issues and PRs welcome. For display bugs, open an issue with a short
repro (what you did, what you expected, what happened), a screenshot of
the panel, and your <code>doctor.log</code>.
</p>

<h2>Credits</h2>
<ul>
  <li>Created by <strong>Asaf Abramzon</strong> - <a href="https://www.linkedin.com/in/asaf-abramzon-7a2b61180/">LinkedIn</a> · <a href="https://github.com/asaf-aizone">GitHub</a>.</li>
  <li><a href="https://github.com/cyrus-and/chrome-remote-interface"><code>chrome-remote-interface</code></a> - CDP client.</li>
</ul>

<h2>Disclaimer</h2>
<p>
Independent open-source tool. Not affiliated with, endorsed by, or connected
to Anthropic or Microsoft. "Claude" is a trademark of Anthropic, PBC.
"Microsoft", "Word", "Excel", and "PowerPoint" are trademarks of Microsoft Corporation. This project
does not redistribute, modify, or contain proprietary code from either company.
</p>

<h3>What this tool does NOT do</h3>

<ul>
  <li>Bypass guardrails, rate limits, or safety systems.</li>
  <li>Reverse-engineer Anthropic's Services, API, or model.</li>
  <li>Scrape or harvest data from Claude or the conversation.</li>
  <li>Provide automated access to Claude. The user drives every conversation manually.</li>
  <li>Bypass Microsoft's add-in security model. The Office add-in is run unmodified by Word, Excel, or PowerPoint exactly as Anthropic ships it.</li>
</ul>

<h2>Further reading</h2>

<ul>
  <li><a href="CHANGELOG.md"><code>CHANGELOG.md</code></a> - changelog per release.</li>
  <li><a href="README.he.md"><code>README.he.md</code></a> - concise Hebrew-only version (text only, no images).</li>
  <li><a href="SECURITY.md"><code>SECURITY.md</code></a> - vulnerability reporting policy.</li>
  <li><a href="docs/security.md"><code>docs/security.md</code></a> - full threat model.</li>
</ul>

<h2>License</h2>
<p>Apache License 2.0 - see <a href="LICENSE"><code>LICENSE</code></a>.</p>

</details>
