# חיבור Google Apps Script

הפרויקט מחובר ל־Apps Script באמצעות `clasp`. קוד הסקריפט נשמר בתיקייה `apps-script`.

## התקנה ראשונה במחשב

1. התקן Node.js בגרסת LTS.
2. פתח PowerShell בתוך תיקיית הריפוזיטורי.
3. הרץ:

```powershell
npm install
npm run gas:login
npm run gas:pull
npm run gas:status
```

בפקודת ההתחברות ייפתח חלון Google. יש לבחור את החשבון שבבעלותו פרויקט ה־Apps Script ולאשר גישה.

## עבודה שוטפת

משיכת הגרסה הנוכחית מגוגל:

```powershell
npm run gas:pull
```

העלאת שינויים ל־Apps Script:

```powershell
npm run gas:push
```

פתיחת הפרויקט ב־Apps Script:

```powershell
npm run gas:open
```

## אבטחה

אין להעלות ל־GitHub את הקובץ `.clasprc.json`, סיסמאות, קודי אימות או מפתחות API. הקובץ `.gitignore` כבר מונע העלאה רגילה של פרטי ההתחברות.
