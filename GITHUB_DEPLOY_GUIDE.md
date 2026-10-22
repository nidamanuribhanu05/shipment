# Publish Shipflow Excel Website on GitHub Pages

This package is the **public demo website**, ready for GitHub Pages. It contains only static HTML, CSS and JavaScript, a sample workbook and instructions. It does not contain the Microsoft 365 backend, credentials or secrets.

## 1. Create a GitHub repository

1. Sign in at https://github.com.
2. Select **New repository** (or visit https://github.com/new).
3. Name the repository `shipflow-demo` (or your preferred name).
4. Select **Public** for a simple GitHub Pages demo. Remember that GitHub Pages sites are public on the internet even when supported accounts publish from a private repository.
5. Click **Create repository**.

## 2. Upload the app files

1. Download `shipflow_github_demo.zip` and **extract it** on your computer.
2. Open your new repository and select **Add file > Upload files**.
3. Drag the **contents of the extracted folder** into the repository, not the ZIP archive itself. The repository root must contain `index.html`, `app.js`, `workbook.js`, `styles.css`, `.nojekyll` and the `samples` folder.
4. Select **Commit changes**. If your browser doesn't include hidden `.nojekyll`, create an empty file named `.nojekyll` in the repository through **Add file > Create new file**. It's optional for most projects but ensures the files are published as-is.

## 3. Turn on GitHub Pages

1. In the repository, select **Settings**.
2. Select **Pages** on the left.
3. Under **Build and deployment**, choose **Deploy from a branch**.
4. Select branch **main**, folder **/(root)**, and click **Save**.
5. Wait for GitHub to publish your site, then select **Visit site** under Pages.

Your demo URL will look like:

`https://YOUR-GITHUB-USERNAME.github.io/shipflow-demo/`

If you used a different repo name, the final part of the URL changes.

## 4. Test the Excel website

1. Click **Excel workspace** in the left sidebar.
2. Choose **Sample data** or **Open Excel** and select an `.xlsx` or `.csv` workbook under 2.5 MB.
3. Double click a cell to edit. You can use arrow keys, paste cells copied from Excel, and add a row or column.
4. Choose **Download .xlsx** to download your edited first worksheet.
5. Choose **Use in dispatch** to put the generated workbook in the shipment form.
6. Enter a **test** recipient and subject, then choose **Simulate dispatch**.
7. Switch **View as** between Admin and User to see the demo screens.

## What works, what does not

| Feature | GitHub demo | Later Microsoft Azure deployment |
| --- | --- | --- |
| Upload `.xlsx` or `.csv` into editor | Yes, in browser | Yes |
| Edit cells, paste, add rows, download `.xlsx` | Yes, in browser | Yes |
| Attach an edited workbook to dispatch form | Yes | Yes |
| Send emails to real addresses | **No, demo only** | Requires Microsoft Graph, work email, authorization |
| Save files to SharePoint | **No** | Requires SharePoint site/library and Graph permissions |
| Enforced admin/user security | **No, UI simulation** | Requires Microsoft Entra ID app roles and backend checks |
| Record history | Sample plus demo metadata saved in this browser | Server-backed audit records |

## Limitations and safety

- **Do not upload confidential or real customer data to a publicly accessible demo**, and do not use it for business workflows. The public website serves client-side code and cannot guarantee tamper-proof access control. Excel file contents are processed locally in the browser; only sample/demo record metadata is stored by the demo in local browser storage.
- No emails, shipment bookings, delivery tracking or SharePoint uploads occur.
- The spreadsheet editor supports the **first worksheet**, reading common `.xlsx` data and CSV. Editing and exporting **does not preserve** advanced Excel formatting, formulas, merged cells, multiple sheets, macros, charts, or complex workbook features. Exported cells contain values. Limit: up to 1,200 rows and 35 columns shown by the demo editor, and 2.5 MB compressed upload file size. Export has an Excel-compatible `.xlsx` ZIP structure.
- The editor needs a recent browser supporting `DecompressionStream('deflate-raw')` to read compressed Excel workbooks. Modern Chrome or Edge is recommended.
- GitHub Pages is a static host. It cannot run the Node.js/Express backend, securely keep Microsoft 365 app secrets, send real mail or enforce actual authorization. Never add `.env`, API keys, client secrets, or private workbooks to this repo.

## After approving the demo

Use the separate `shipflow_microsoft_later.zip` package. Its `public/` folder has the same Excel interface, while `server.js` provides the Microsoft 365 integration starter. Follow `AZURE_DEPLOY.md` only after adding production security protections such as durable encrypted sessions, a persistent token cache, malware scanning and an independent security review.

Documentation: https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site
