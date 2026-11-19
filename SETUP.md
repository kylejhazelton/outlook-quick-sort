# Quick Sort: Outlook add-in setup

A pinnable side panel for new Outlook (Windows and web) that shows every category in your mailbox as a colored chip. Click a chip to add that category to the open email, and click it again to remove it. A checkmark shows which categories are already applied, and the panel follows you as you go from email to email.

It needs no sign-in and no app registration. It only uses Outlook's built-in add-in API.

## Files

| File | What it is |
|---|---|
| `manifest.xml` | Tells Outlook where the add-in lives. This is the file you sideload. |
| `taskpane.html / .js / .css` | The panel itself. |
| `assets/` | Icons. |
| `set-host.ps1` | Rewrites the URLs in `manifest.xml` to point at your host. |

Open `taskpane.html` outside Outlook and it runs in **preview mode** with sample data.

---

## Step 1: Host the files over HTTPS

Outlook loads the panel from a URL, so the files have to be on an HTTPS host that's always reachable. **GitHub Pages** is free and the easiest option:

1. Create a free GitHub account if you don't have one, then create a new repo, e.g. `outlook-quick-sort`.
2. Click **Add file → Upload files** and drag in everything in this folder, including the `assets` folder. Then click **Commit changes**.
3. Go to repo **Settings → Pages**, set Source to *Deploy from a branch*, choose `main` and `/ (root)`, then click **Save**.
4. After a minute, check that `https://<your-username>.github.io/outlook-quick-sort/taskpane.html` opens and shows the preview.
5. In PowerShell, from this folder:

   ```powershell
   .\set-host.ps1 -BaseUrl https://<your-username>.github.io/outlook-quick-sort
   ```

   This updates `manifest.xml` on your PC. You don't need to re-upload it, because Outlook reads the manifest from the file you sideload.

## Step 2: Sideload into Outlook

1. In new Outlook (or Outlook on the web, since the same add-in list applies to both), open <https://aka.ms/olksideload>. Or go to **Apps / Get Add-ins → My add-ins**.
2. Under **Custom add-ins**, choose **Add a custom add-in → Add from file…** and pick `manifest.xml`.
3. Open any email. **Quick Sort** appears in the message's **Apps** button (or under the **…** menu).
4. Click the **pin** icon at the top of the panel so it stays open as you move between emails.

## Moving emails in one click (no add-in needed)

Use Outlook's built-in **Quick Steps** (Home tab → **Quick steps → Manage quick steps**). Make one per folder you file into often, using the action *Move to folder*. You can also add *Categorize* to the same step and give it a shortcut like Ctrl+Shift+1.

## Updating later

- **Panel changes** (HTML/JS/CSS): re-upload the files to GitHub. Outlook picks them up the next time the panel opens. You may need to close and reopen Outlook to clear its cache.
- **Manifest changes:** remove the add-in in *My add-ins* and add it again.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Panel is blank | Open the `taskpane.html` URL in a browser to confirm it's reachable over HTTPS. |
| New category doesn't show up | Click ↻ in the panel. |
| "Add from file" option is missing | Your organization may have turned off custom add-ins. Try Outlook on the web at <https://aka.ms/olksideload>. If it's blocked there too, IT has disabled sideloading. |
