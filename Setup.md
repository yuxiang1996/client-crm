# Client CRM: how it fits together

- **The app:** `index.html`, hosted on GitHub Pages at `https://yuxiang1996.github.io/client-crm/`.
- **The data:** your **Master List** sheet, tab **CRM_Clients**. The app reads and writes it through the Apps Script project **Client CRM** in kennkwangyx@gmail.com (`Code.gs`).
- **The lock:** the Apps Script only answers when the request carries the access key, which is stored in its Script Properties as `CRM_KEY`. The key is never in this repo. Each phone or computer gets it once from your setup link, and it stays in that browser.

## Putting it on a phone

1. Open your **setup link** (it's in `Client CRM access key.txt` in your Claude outputs folder, not in GitHub).
2. Tap **Share**, then **Add to Home Screen**.

Lost phone? In Apps Script, go to **Project Settings**, then **Script Properties**, change `CRM_KEY` to something new, and make a new setup link. Every old device is locked out at once.

## How it uses your sheets

- **Done / Log** sets **Last Followed Up** to today. Your **Next Follow Up** formula works out the next date, and the app never writes over a formula cell.
- Notes go into the **CRM_Log** tab. If "Also add to Remarks" is ticked, they're also added to **Other Remarks** as `19/11: note`.
- **+ (Add client)** fills the first row with a blank Name, reusing your pre-made Client IDs.
- **Import PPC leads** (bottom of the Clients screen) copies rows from **Our Master List → CRM** whose **PIC** is Kwang or Seah. It skips anyone already in your list (matched by phone) and Topic = FAKE. It only reads Our Master List, never changes it. PIC, project and category go into Remarks, so searching "Seah" or "Thomson" finds them.

## Changing the code

- **Page (`index.html`):** commit and push to GitHub. Pages updates in about a minute.
- **Apps Script (`Code.gs`):** paste into the project, press **Ctrl + S**, then **Deploy**, **Manage deployments**, the pencil icon, **Version: New version**, **Deploy**. The URL stays the same.
