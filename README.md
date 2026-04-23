<div dir="rtl">

<h1>Claude for Word RTL Fix</h1>

<p>
תיקון RTL לפאנל של Claude ב-Microsoft Word. מקומי, בלי טלמטריה, Apache 2.0.<br>
<em>לכל מי שכותב עברית ב-Word עם תוסף Claude. Windows 10/11.</em>
</p>

<p>
<strong>אינו תוסף רשמי של Anthropic או של Microsoft.</strong> כלי open-source עצמאי.
</p>

<blockquote>
<p>
<strong>Windows בלבד.</strong> הכלי לא עובד על macOS או Linux. תוסף Claude ל-Word
מבוסס על WebView2 של מיקרוסופט, שקיים רק ב-Windows. ל-Word ל-Mac יש runtime
אחר (WKWebView) שלא חושף את אותו debugging interface, וכל שכבת ההפעלה (bat, vbs,
PowerShell, Registry, Startup folder) לא רלוונטית שם. אם אתם על Mac, אין port מ-Word.
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
הפאנל של Claude בתוך Word לא תומך ב-RTL. עברית יוצאת הפוכה, bullets
בצד הלא נכון, פיסוק נופל איפה שלא צריך. הכלי מתחבר ל-WebView2 של
הפאנל דרך Chrome DevTools Protocol, מזריק CSS ו-MutationObserver, והפאנל
עובר ל-RTL תקין.
</p>

<p>
המודל של Claude, ה-API של Anthropic, ו-Word עצמו לא נגועים. רק ה-DOM
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
אייקון קטן ליד השעון: ריבוע מעוגל עם האות <strong>W</strong> וחץ RTL
לבנים, וצבע רקע שמשקף את המצב. אפור = בטעינה, אדום = מנותק, ירוק =
RTL פעיל. קליק ימני פותח תפריט קצר. אין כניסה ב-Start Menu, אין שינוי
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
      <td>Connect (relaunch Claude for Word RTL Fix)</td>
      <td>האייקון אדום</td>
      <td>מעלה מחדש את Word דרך ה-wrapper עם debug-port פתוח. אם Word רץ - שומר מסמכים, מבקש אישור, מריץ מחדש עם אותם מסמכים</td>
    </tr>
    <tr>
      <td>Disconnect (close Claude for Word RTL Fix)</td>
      <td>תמיד</td>
      <td>כפתור התאוששות אוניברסלי. עוצר timers של Connect באוויר, סוגר את Word (graceful + force כגיבוי), הורג את ה-injector, מנקה את כל ה-state files. אם משהו השתבש - לחיצה כאן תמיד מחזירה למצב אפס</td>
    </tr>
    <tr>
      <td>Auto-enable at every Word launch (checkbox)</td>
      <td>תמיד</td>
      <td>מדליק RTL לכל פתיחה עתידית של Word בלי שתצטרכו ללחוץ Connect. ראו "שני מצבי עבודה" למטה</td>
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

<h3>שני מצבי עבודה</h3>

<table dir="rtl">
  <thead>
    <tr><th>מצב</th><th>מה קורה</th><th>למי מתאים</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>ידני (ברירת מחדל)</td>
      <td>כל פעם שפותחים את Word, האייקון אדום. לחיצה על <strong>Connect</strong> מריצה את Word מחדש דרך ה-wrapper. RTL פעיל רק כשבוחרים בכך</td>
      <td>למי שרוצה שליטה מלאה, או עובד לעיתים רחוקות עם Claude בעברית</td>
    </tr>
    <tr>
      <td>אוטומטי (Auto-enable)</td>
      <td>סימון ה-checkbox בתפריט מוסיף משתנה סביבה קבוע ל-user profile. מעכשיו <strong>כל</strong> פתיחה של Word מתחילה עם RTL מופעל. set-it-and-forget-it</td>
      <td>למי שעובד קבוע עם Claude בעברית ורוצה שהכל יעבוד אוטומטית</td>
    </tr>
  </tbody>
</table>

<p>
<strong>גילוי נאות על Auto-enable:</strong> המשתנה שנוסף הוא
<code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code> ברמת <code>HKCU\Environment</code>.
WebView2 הוא רכיב משותף, ולכן המשתנה נקרא <em>גם</em> על ידי אפליקציות
WebView2 אחרות שפועלות תחת אותו user - כמו Teams, Outlook (new), וחלונות
Edge WebView אחרים. בפועל אף אחת מהן לא מאזינה ל-debug port על
<code>9222</code>, אבל המשתנה כן נקרא על ידן וזה משהו שנכון לדעת. אם
לא מקובל עליכם - השאירו את ה-Auto-enable כבוי והשתמשו ב-Connect ידני.
את ה-checkbox אפשר לכבות בכל רגע מהתפריט והמשתנה ינוקה.
</p>

<p>
ב-uninstall, ה-Auto-enable מנוקה אוטומטית - אבל רק אם הערך של המשתנה
תואם בדיוק לזה שלנו. ערך שמשתמש הוסיף ידנית לא נמחק.
</p>

<p>
  <img src="docs/images/auto-enable-on-dialog.png" alt="דיאלוג הפעלת Auto-enable" width="420">
  <img src="docs/images/auto-enable-off-dialog.png" alt="דיאלוג כיבוי Auto-enable" width="420">
  <br>
  <em>שני כיווני ה-toggle - הפעלה (מימין) וכיבוי (משמאל). הטקסט מסביר בכל כיוון מה ישתנה.</em>
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
מנקה את משתנה הסביבה של Auto-enable (אם הוא תואם לערך שלנו), ומפנק
את התלויות. המטרה: להסיר את כל מה שההתקנה יצרה. ערכים שהמשתמש הגדיר
ידנית באותו משתנה סביבה לא נמחקים.
</p>

<hr>

<h2 id="install">התקנה</h2>

<p>
<strong>דרישות:</strong> Windows 10/11, Microsoft Word desktop עם תוסף Claude מותקן.
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
  <li>להוריד זיפ מ-<a href="https://github.com/asaf-aizone/Claude-for-word-RTL-fix/releases">Releases</a>, או <code>git clone</code>.</li>
  <li>לחלץ לתיקייה שתישמר (למשל <code>C:\Tools\claude-word-rtl\</code>).</li>
  <li>לסגור את Word אם הוא פתוח.</li>
  <li>דאבל-קליק על <strong><code>install.bat</code></strong>. בסיום הוא ישאל אם להפעיל <em>Auto-enable</em> - מומלץ Y.</li>
  <li>אייקון ה-tray יעלה ליד השעון, ויעלה אוטומטית גם בלוגין הבא.</li>
</ol>

<p>
לוגים נכתבים ל-<code>install.log</code> ליד ה-installer. לא נדרשות הרשאות admin.
</p>

<p>
  <img src="docs/images/installer-done.png" alt="פלט install.bat בסיום מוצלח - Auto-enable מופעל, Stopped previous tray, Installation complete, ו-How to use" width="640">
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

<p>
אחרי התקנה, האייקון עולה אוטומטית בכל לוגין.
</p>

<p><strong>מצב ידני (ברירת מחדל):</strong></p>

<ol>
  <li>פותחים את Word איך שרגילים - Start Menu, דאבל-קליק על <code>.docx</code>, Recent files.</li>
  <li>אם האייקון <strong>אדום</strong> (ה-injector לא מחובר לפאנל): קליק ימני, <strong>Connect</strong>.</li>
  <li>דיאלוג יקפוץ עם הסבר על מה שיקרה. אם יש מסמכים לא שמורים - תופיע אזהרה מפורשת. אישור סוגר את Word ומריץ אותו מחדש עם אותם מסמכים.</li>
  <li>האייקון נהפך ל<strong>ירוק</strong>. הפאנל RTL.</li>
</ol>

<p>
  <img src="docs/images/connect-dialog.png" alt="דיאלוג האישור של Connect - עם אזהרת UNSAVED" width="500">
</p>

<p><strong>מצב אוטומטי (Auto-enable מסומן):</strong></p>

<ol>
  <li>פותחים את Word איך שרגילים. זהו.</li>
  <li>האייקון בתפריט הופך לירוק תוך שניות, הפאנל כבר RTL.</li>
  <li>בפועל אין צורך לפתוח את התפריט - האייקון משמש כאינדיקטור סטטוס בלבד.</li>
</ol>

<hr>

<h2>איך זה עובד?</h2>

<p>
<code>word-wrapper.bat</code> מריץ את Word עם משתנה סביבה
<code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222</code>.
זה flag רשמי של Microsoft שפותח Chrome DevTools Protocol ב-<code>localhost:9222</code>.
תהליך Node קטן (<code>scripts/inject.js</code>) מתחבר דרך WebSocket מקומי,
מאתר את ה-page שה-URL שלו תואם לפאנל של Claude, ומריץ <code>Runtime.evaluate</code>
כדי להזריק <code>&lt;style&gt;</code> ו-MutationObserver. לולאה של שתי שניות מזריקה מחדש
אם הפאנל טוען את עצמו.
</p>

<p>
ה-injector לא יוצר ולא שולח בקשות HTTP. הפאנל ממשיך לדבר ישירות מול
Anthropic כמו תמיד - ה-wrapper פשוט מוסיף flag ל-WebView2 ומחכה בצד.
</p>

<p>
ה-Connect לא תוקע את ה-UI: התפריט נסגר מיד, וה-state machine ממשיך ברקע
דרך timer. אם Word לא נסגר תוך 10 שניות, מופיע דיאלוג OK/Cancel - OK
מחסל את התהליך ומריץ מחדש, Cancel משאיר את Word כפי שהוא.
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
      <td>WebView2 של Word</td>
      <td>קריאה דרך Chrome DevTools Protocol על <code>localhost:9222</code></td>
      <td>לאתר את הפאנל של Claude ולהזריק CSS</td>
    </tr>
    <tr>
      <td>משתנה סביבה <code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code></td>
      <td>במצב ידני: כתיבה בהקשר ה-wrapper בלבד. במצב Auto-enable: כתיבה קבועה ב-<code>HKCU\Environment</code></td>
      <td>לפתוח את ה-debug port ב-WebView2 ברגע הפעלה של Word. במצב Auto-enable - לכל פתיחה עתידית</td>
    </tr>
    <tr>
      <td><code>%TEMP%</code></td>
      <td>כתיבה של PID ו-status של ה-injector</td>
      <td>למעקב אחר מצב ההזרקה מה-tray ומניעת mass-kill של תהליכי Node</td>
    </tr>
    <tr>
      <td>Startup folder של המשתמש</td>
      <td>יצירת קיצור אחד לכניסה</td>
      <td>להפעיל את ה-tray אוטומטית בלוגין</td>
    </tr>
    <tr>
      <td><code>HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL</code></td>
      <td>כתיבה בהתקנה, מחיקה בהסרה</td>
      <td>רישום הכלי ב-Windows Settings &gt; Apps כדי שיופיע ברשימת Installed apps ויהיה ניתן להסרה משם</td>
    </tr>
    <tr>
      <td>Microsoft Word (COM)</td>
      <td>קריאה בלבד, רק בזמן Connect</td>
      <td>למנות מסמכים פתוחים כדי לפתוח אותם מחדש אחרי relaunch</td>
    </tr>
  </tbody>
</table>

<p>
<strong>בנוסף לטבלה למעלה, הכלי לא נוגע ב:</strong> file associations,
<code>Normal.dotm</code>, תוספים אחרים של Word, או services.
ב-registry הוא נוגע רק בשני המפתחות שבטבלה
(<code>HKCU\...\Uninstall\ClaudeWordRTL</code> תמיד,
ו-<code>HKCU\Environment</code> רק כש-Auto-enable פעיל).
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
כל עוד Word רץ דרך הכלי, WebView2 פותח debug port ב-<code>localhost:9222</code>.
ה-port לא חשוף לרשת, אבל תהליך מקומי באותו user יכול להתחבר אליו (זה
המנגנון הסטנדרטי של Chrome DevTools Protocol; כל דפדפן מבוסס Chromium
שפותח debug port מתנהג אותו דבר). בפועל, כמעט כל מה שרץ בסשן שלך יכול
כבר לקרוא את הזיכרון של Word. השימוש ב-<strong>Disconnect</strong> מה-tray
כשסיימתם מנקה את ה-port. על מחשבים משותפים או לא מהימנים - אל תריצו.
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

<p><strong>מה הכלי עושה ל-Word?</strong><br>
לא משנה את Word עצמו: בלי patch לתוסף, בלי שינוי טמפלטים, בלי
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

<p><strong>Word Online? Mac? Microsoft 365 תאגידי עם EDR?</strong><br>
Word Online - לא, הכלי דורש WebView2 של Word desktop. Mac - לא, Windows
בלבד. תאגידי - לבדוק עם IT לפני הפעלה של debug port ב-Office. לא
מיועד ל-laptops תאגידיים סגורים.</p>

<p><strong>Word פתוח עם 10 מסמכים. Connect ייסגור אותם?</strong><br>
כן, אבל בעדינות - Word מתבקש לשמור שינויים, מקבל את רשימת המסמכים דרך
COM, וה-wrapper פותח את כולם מחדש. אם משהו לא נשמר, Word ישאל כרגיל.</p>

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
      <td>Auto-enable לא לוקח אפקט באפליקציה אחרת (Teams/Outlook)</td>
      <td>צפוי. משתני סביבה נטענים בפתיחת תהליך. צריך לסגור ולפתוח מחדש את האפליקציה. שינוי משודר דרך <code>WM_SETTINGCHANGE</code> אבל לא משפיע על תהליכים שכבר רצים</td>
    </tr>
    <tr>
      <td>הטריי נשאר אדום למרות ש-Auto-enable דלוק, Word פתוח, ו-Node מותקן</td>
      <td>פורט 9222 אולי תפוס בידי אפליקציה אחרת. בודקים ב-cmd: <code>netstat -ano | findstr :9222</code>. אם רואים שורה עם PID שאינו של <code>WINWORD.EXE</code>, אפליקציה אחרת יושבת על הפורט. אשמים מוכרים: Google Drive File Stream, אפליקציות מבוססות Electron שמריצות עם <code>--remote-debugging-port=9222</code>, או WebView2 SDK tools. סוגרים את האפליקציה שתופסת את הפורט (לפי ה-PID וה-process name ב-Task Manager) ואז פותחים את Word מחדש דרך הטריי (Connect). מגרסה 0.1.3 ואילך ה-injector מטפל במקרה שבו גם Drive וגם Word יושבים על 9222 בו-זמנית בגלל IPv4/IPv6 split, אבל אם אף אחד לא מציע panel של Claude, אין מה לתפוס. <code>doctor.bat</code> של גרסה 0.1.3 מציג את זה בבירור בשתי הבדיקות החדשות.</td>
    </tr>
  </tbody>
</table>

<hr>

<h2>מגבלות ידועות</h2>

<ul>
  <li>debug port ב-<code>localhost:9222</code> לא מאומת - כל תהליך מקומי באותו user יכול להתחבר. ראו "הערת אבטחה".</li>
  <li>Microsoft 365 תאגידי עם EDR/DLP יכול לחסום את דגל ה-WebView2. הכלי לא מיועד ללפטופים ארגוניים סגורים.</li>
  <li>עדכון של תוסף Claude שמחליף את ה-DOM יכול לשבור את ההזרקה עד patch. יישלח release מתוקן.</li>
  <li><strong>Mac (macOS) לא נתמך ולא יהיה נתמך.</strong> Word ל-Mac משתמש ב-WKWebView במקום WebView2, ושכבת ה-launcher כולה (bat, vbs, ps1) היא Windows-only. Word Online גם לא נתמך.</li>
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
ולא מכיל קוד שלהן. "Claude" סימן מסחרי של Anthropic, PBC. "Microsoft"
ו-"Word" סימנים מסחריים של Microsoft Corporation.
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

<h1>Claude for Word RTL Fix</h1>

<p>
RTL fix for the Claude panel in Microsoft Word. Local-only, no telemetry, Apache 2.0.
</p>

<p>
<strong>Not an official Anthropic or Microsoft add-in.</strong> Independent open-source tool.
</p>

<blockquote>
<p>
<strong>Windows only.</strong> This tool does not work on macOS or Linux. The Claude
Word add-in is built on Microsoft's WebView2 runtime, which is Windows-only. Word
for Mac uses WKWebView, which does not expose the same debugging interface, and the
entire launcher stack (batch, VBS, PowerShell, registry, Startup folder) is
Windows-specific. If you are on Mac, this tool has no port.
</p>
</blockquote>

<h2>What it does</h2>

<p>
Anthropic's official Claude add-in for Word doesn't render Hebrew right-to-left.
Bullets land on the wrong side, alignment is reversed, punctuation ends up in
weird places. This tool attaches to the add-in's WebView2 panel via Chrome
DevTools Protocol, injects a small stylesheet plus a MutationObserver, and
flips the panel to correct RTL. The Claude model, the Anthropic API, and Word
itself are untouched - only the panel's local DOM, and only while the panel
is open.
</p>

<h2>Features</h2>

<ul>
  <li><strong>Instant RTL</strong> - direction, alignment, bullets, tables. <code>&lt;pre&gt;</code> and <code>&lt;code&gt;</code> stay LTR so source code isn't corrupted.</li>
  <li><strong>Tray-icon control</strong> - a small rounded-square icon near the clock: a white <strong>W</strong> and an RTL arrow on a status-colored background. Gray = starting. Red = not attached. Green = panel is RTL. Right-click for Connect / Disconnect / Auto-enable / Show diagnostic log / Check for updates / Uninstall / Exit. No Start Menu entry pretending to be Word. See the <a href="docs/images/tray-icon-red.png">red (disconnected)</a> and <a href="docs/images/tray-icon-green.png">green (connected)</a> states.</li>
  <li><strong>Two modes</strong> - Manual (click Connect each time) or Auto-enable (set-it-and-forget-it). See "Auto-enable" below for the trade-off.</li>
  <li><strong>Non-blocking Connect</strong> - the menu closes immediately, work proceeds on a background timer state machine. If Word doesn't close within 10 seconds, an OK/Cancel dialog offers force-kill or abort.</li>
  <li><strong>Minimal footprint</strong> - no file associations, no <code>Normal.dotm</code> changes, no templates, no services. Install creates two per-user items: a Startup-folder shortcut (so the tray auto-launches at login) and an <code>HKCU\...\Uninstall\ClaudeWordRTL</code> registry key (so the tool appears in Windows Settings &gt; Apps). Both are removed by <code>uninstall.bat</code>.</li>
  <li><strong>Hebrew typography cleanup</strong> - em-dash (—) and en-dash (–) become hyphen (-). Arrows (→ ← ↔ ⇒ ⇐) become commas. Input fields are left alone.</li>
  <li><strong>Crash-safe status</strong> - if the injector dies without cleanup, the tray detects stale state (process + status-file timestamp) and flips to red, to reduce cases where the icon shows green without an actual connection.</li>
  <li><strong>Diagnostic log</strong> - <code>%TEMP%\claude-word-rtl.log</code>, accessible from the tray menu. Truncated on each injector start. Shows discovered CDP targets, attach events, errors.</li>
  <li><strong>Clean uninstall</strong> - 4-step <code>uninstall.bat</code>: stop tray/injector, remove Startup entry + <code>HKCU\...\Uninstall\ClaudeWordRTL</code>, clear Auto-enable env var (only if it still matches our exact value - user-modified values are preserved), prune deps. Aims to remove everything the installer created.</li>
</ul>

<h3>Tray menu</h3>

<table>
  <thead>
    <tr><th>Item</th><th>When available</th><th>What it does</th></tr>
  </thead>
  <tbody>
    <tr><td>Connect (relaunch Claude for Word RTL Fix)</td><td>Icon is red</td><td>Relaunches Word through the wrapper with the debug-port enabled. If Word is already open, saves documents, asks for confirmation, and reopens them. Non-blocking - menu closes immediately.</td></tr>
    <tr><td>Disconnect (close Claude for Word RTL Fix)</td><td>Anytime</td><td>Universal recovery button. Stops any in-flight Connect timers, closes Word (graceful + force fallback), kills the injector, cleans up state files. If anything went wrong, clicking this always returns to a clean slate.</td></tr>
    <tr><td>Auto-enable at every Word launch (checkbox)</td><td>Anytime</td><td>Persists <code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code> in <code>HKCU\Environment</code> so every future Word launch starts with RTL. See Auto-enable section for caveats.</td></tr>
    <tr><td>Show diagnostic log</td><td>Anytime</td><td>Opens <code>%TEMP%\claude-word-rtl.log</code> in the default editor.</td></tr>
    <tr><td>Check for updates...</td><td>Anytime</td><td>Runs <code>check-update.js</code> and shows the result in a dialog. When a newer release exists, a one-click button opens the download page in the default browser.</td></tr>
    <tr><td>Uninstall...</td><td>Anytime</td><td>Runs <code>uninstall.bat</code> after confirmation.</td></tr>
    <tr><td>Exit</td><td>Anytime</td><td>Stops the tray icon.</td></tr>
  </tbody>
</table>

<h3>Auto-enable: the trade-off</h3>

<p>
Auto-enable persists the WebView2 debug-port flag at user level. The
variable name is <code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code>, and it's
read by every WebView2 host running under your user - Teams, the new
Outlook, Edge WebView windows, and any other WebView2-based app. In
practice most of them do not actually bind the debug port on
<code>9222</code>, but the variable is still read by them and that's worth
knowing. Uncheck the menu item at any time to remove it.
</p>

<p>
If that's not acceptable, leave Auto-enable off and use manual Connect.
On uninstall, the variable is cleared only if its value still matches
exactly what we wrote - any user-modified value is preserved.
</p>

<h2 id="install-en">Install</h2>

<p>
<strong>Requirements:</strong> Windows 10/11, Microsoft Word desktop with Claude
add-in installed.
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
  <li>Download the zip from <a href="https://github.com/asaf-aizone/Claude-for-word-RTL-fix/releases">Releases</a> or <code>git clone</code>.</li>
  <li>Extract to a folder you'll keep (e.g. <code>C:\Tools\claude-word-rtl\</code>).</li>
  <li>Close Word if it's open.</li>
  <li>Double-click <strong><code>install.bat</code></strong>. At the end it will ask whether to turn on <em>Auto-enable</em> - Y is recommended.</li>
  <li>Tray icon appears near the clock, and will launch automatically on every login.</li>
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
  <li>Close Word.</li>
  <li>Run <code>install.bat</code> again. It stops the old tray and
      injector and starts the new ones with the updated code.</li>
</ol>

<h3>Uninstall</h3>
<p>Double-click <strong><code>uninstall.bat</code></strong>.</p>

<h2>Daily use</h2>

<ol>
  <li>Open Word the way you always do - Start Menu, double-click a <code>.docx</code>, Recent files.</li>
  <li>If the icon is <strong>red</strong>: right-click, <strong>Connect</strong>.</li>
  <li>The tool enumerates your open documents, asks for confirmation, reopens them.</li>
  <li>Icon turns <strong>green</strong>. Panel is RTL.</li>
</ol>

<h2>How it works</h2>

<p>
<code>word-wrapper.bat</code> launches Word with
<code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222</code>,
a Microsoft-documented flag that exposes Chrome DevTools Protocol on
<code>localhost:9222</code>. A small Node process (<code>scripts/inject.js</code>) connects
over the local WebSocket, finds the page whose URL matches the Claude panel,
and calls <code>Runtime.evaluate</code> to inject a <code>&lt;style&gt;</code> element and
a MutationObserver. A 2-second poll re-injects if the panel reloads.
</p>

<p>
All activity is local. The only outbound traffic is normal Claude-to-Anthropic
traffic, unchanged, routed by Word itself.
</p>

<p>Full threat model: <a href="docs/security.md"><code>docs/security.md</code></a>.</p>

<h2>What the tool accesses</h2>

<table>
  <thead>
    <tr><th>Resource</th><th>Access</th><th>Why</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>Word's WebView2</td>
      <td>Read via Chrome DevTools Protocol on <code>localhost:9222</code></td>
      <td>Locate the Claude panel and inject CSS</td>
    </tr>
    <tr>
      <td><code>WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS</code> env var</td>
      <td>Manual mode: written in the wrapper's context only. Auto-enable mode: persisted in <code>HKCU\Environment</code></td>
      <td>Open the debug port at Word launch (or for every future WebView2 host launch when Auto-enable is on)</td>
    </tr>
    <tr>
      <td><code>%TEMP%</code></td>
      <td>Writes injector PID and status</td>
      <td>Track injection state from the tray, prevent mass-kill of Node processes</td>
    </tr>
    <tr>
      <td>User Startup folder</td>
      <td>Creates one shortcut</td>
      <td>Launch the tray at login</td>
    </tr>
    <tr>
      <td><code>HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL</code></td>
      <td>Written on install, removed on uninstall</td>
      <td>Register the tool in Windows Settings &gt; Apps &gt; Installed apps so it can be uninstalled from there</td>
    </tr>
    <tr>
      <td>Microsoft Word (COM)</td>
      <td>Read-only, only during Connect</td>
      <td>Enumerate open documents to reopen them after relaunch</td>
    </tr>
  </tbody>
</table>

<p>
<strong>Beyond the table above, the tool does not touch:</strong> file
associations, <code>Normal.dotm</code>, other Word add-ins, or Windows
services. The only registry keys it writes are the two in the table
(<code>HKCU\...\Uninstall\ClaudeWordRTL</code> always, and
<code>HKCU\Environment</code> only when Auto-enable is on).
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
While Word runs via this tool, WebView2 opens a debug port on
<code>localhost:9222</code>. Any local process on the same user can connect and
read the panel's DOM. The port is localhost-only, but unauthenticated.
Use <strong>Disconnect</strong> when done. Don't run on shared or untrusted machines.
</p>

<h2>FAQ</h2>

<p><strong>Will Anthropic block me?</strong><br>
Not expected. The tool only changes how the panel renders in your local DOM.
It doesn't alter what you send, what Claude returns, rate limits, or
guardrails. That said, Anthropic's terms of service are what matter and
responsibility for usage is yours.</p>

<p><strong>Does it modify Word?</strong><br>
It doesn't modify Word itself: no template changes, no add-in patches, no
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

<p><strong>Word Online, Mac, corporate M365 with EDR?</strong><br>
Word Online - no, requires Word desktop's WebView2. Mac - no, Windows only.
Corporate - check with your IT team before enabling a WebView2 debug port.
Not intended for sealed corporate laptops.</p>

<p><strong>Word is open with 10 documents. Will Connect close them?</strong><br>
Yes, gracefully. Word is asked to save changes, the document list is captured
via COM, and the wrapper reopens all of them. If something wasn't saved,
Word will prompt as usual.</p>

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
      <td>Auto-enable doesn't take effect in Teams/Outlook</td>
      <td>Expected. Env vars load at process start. Close and reopen the app. <code>WM_SETTINGCHANGE</code> is broadcast but doesn't affect already-running processes</td>
    </tr>
    <tr>
      <td>Tray stays red even with Auto-enable on, Word open, and Node installed</td>
      <td>Port 9222 may be owned by another app. In cmd: <code>netstat -ano | findstr :9222</code>. If you see a row with a PID that is not <code>WINWORD.EXE</code>, another app is squatting on the port. Known offenders: Google Drive File Stream, Electron-based apps launched with <code>--remote-debugging-port=9222</code>, WebView2 SDK tools. Close the squatting app (find its PID in Task Manager), then reopen Word via tray, Connect. As of v0.1.3 the injector handles the common case where Drive and Word each bind a different family (IPv4 vs IPv6) of <code>localhost:9222</code>; but if Word is not actually exposing the Claude panel on 9222, there is nothing to attach to. <code>doctor.bat</code> as of v0.1.3 surfaces this directly in its two new checks.</td>
    </tr>
  </tbody>
</table>

<h2>Known limitations</h2>

<ul>
  <li>The <code>localhost:9222</code> debug port is unauthenticated. See Security note.</li>
  <li>Corporate M365 with EDR/DLP may block the WebView2 flag. Not intended for sealed corporate laptops.</li>
  <li>A Claude add-in update that changes the panel DOM can break injection until a patch is released.</li>
  <li><strong>Mac (macOS) is not supported and will not be supported.</strong> Word for Mac uses WKWebView instead of WebView2, and the entire launcher stack (bat, vbs, ps1) is Windows-only. Word Online is also not supported.</li>
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
"Microsoft" and "Word" are trademarks of Microsoft Corporation. This project
does not redistribute, modify, or contain proprietary code from either company.
</p>

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
