# Quick Sort: Outlook add-in

A pinnable side panel for new Outlook (Windows and web) for categorizing email in one click. It needs no sign-in or app registration, because it only uses Outlook's built-in add-in API.

- **Top 5:** your most-used categories, weighted so one-off picks don't jump in (see *How Top 5 works*). Shown alphabetically so they don't move around.
- **All:** every category in your mailbox, alphabetical.
- **One click** adds a category to the open email, and clicking again removes it.
- **Focus stays in your email list,** so you can arrow through emails and click categories without clicking back.
- **Multi-select:** select several emails (same folder) and the panel shows "N emails selected".
  - **✓** means all of them have the category, **–** (dashed outline) means some do.
  - A click adds the category to the ones missing it. If all of them have it, the click removes it from all.

**Where things live**

- **Live add-in:** <https://kylejhazelton.github.io/outlook-quick-sort/> (GitHub repo `kylejhazelton/outlook-quick-sort`)
- **Working copy:** this folder (OneDrive → `_KH Plugins\outlook-quick-sort`)
- **Top 5 scores:** saved in your mailbox, so they follow you between desktop and web and survive a reinstall

---

## Updating the add-in

### Panel changes (`taskpane.html`, `taskpane.js`, `taskpane.css`)

1. On GitHub, open the `outlook-quick-sort` repo → **Add file → Upload files**.
2. Drag in the changed files from this folder → **Commit changes**.
3. Wait about a minute for GitHub Pages to update.
4. Close and reopen the Quick Sort panel. If it still looks old, restart Outlook (see *Restarting Outlook properly* below).

No reinstall is needed.

### Manifest changes (`manifest.xml`): reinstall

Upload `manifest.xml` to GitHub along with the other files, then reinstall:

1. **Open the add-ins window.** In a web browser, go to **<https://aka.ms/olksideload>** and sign in with your work account if asked.
   - It opens Outlook on the web with the **Add-ins for Outlook** window.
   - Desktop Outlook's **Apps** store has no "My add-ins" option, and no menu leads to this window anymore. **The link is the only way in.** Bookmark it.
2. **Remove the old version.** Click **My add-ins** on the left → scroll to **Custom Addins** → on the Quick Sort card click **…** → **Remove**.
3. **Add the new version.** In **Custom Addins**, click **+ Add a custom add-in → Add from file…** → pick `manifest.xml` from this folder → **Install**. The card should show **✓ Added**.
4. **Restart desktop Outlook properly** (see below).
5. **Give it a minute.** After a reinstall, desktop Outlook can take a minute or two before Quick Sort shows up. That's normal.
6. **Pin it to the ribbon again:** see *Pinning to the ribbon* below.

### Restarting Outlook properly

Closing the window isn't always enough. Check the system tray by the clock, and if Outlook's icon is there, right-click it → **Quit**. Then reopen Outlook.

### Pinning to the ribbon

1. Open any email.
2. On the **Home** ribbon, in the add-ins area, click **More apps**.
3. Find **Quick Sort**, click its **…** (or right-click it) → **Pin**.
4. Open Quick Sort on an email and click the **pin icon at the top of the panel**, so it stays open as you move between emails.

*Settings → Mail → Customize actions* does **not** work for this in our version of Outlook. It has no Toolbar section.

---

## How Top 5 works

- **Scoring:** each time you add a category, every score is multiplied by **0.95** (a slow fade). The chosen category gets **+1 per email** it was added to: +1 for one email, +12 for twelve selected emails.
- **Membership:** the five highest scores make up Top 5. A newcomer only replaces the weakest member if it beats it by **0.5**, so close categories don't flip back and forth.
- **What doesn't count:** removing a category isn't a pick.
- **Tuning:** the constants are at the top of `taskpane.js`:
  - `DECAY = 0.95`: lower means it reacts faster to new habits, higher means steadier.
  - `MARGIN = 0.5`: how much a newcomer has to beat the weakest member by.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Quick Sort isn't in More apps (even when searching) | Desktop hasn't loaded it yet. Wait a minute and restart Outlook properly. To check the install itself, open an email at <https://outlook.office.com> and look under the **Apps** grid icon in the email header. |
| Panel looks like an old version | Restart Outlook properly. Outlook caches the panel files. |
| Panel is blank | Open the GitHub Pages link above in a browser. It should show "Preview mode: sample data". |
| New category doesn't show up | Click ↻ in the panel. |
| Multi-select shows a single subject | The selected emails must be in the same folder, unless you use Conversation view. Maximum 100 emails at a time. |
| Arrow keys stop working after clicking | Click once in the email list. Clicking in the panel's empty space (not a category) can still take the focus. |

## Previewing without Outlook

Open the GitHub Pages link in a browser to see the panel with sample data. Add `?multi=4` to the end of the address to preview a 4-email selection.

## Moving emails in one click (no add-in needed)

Use Outlook's built-in **Quick Steps** (Home → **Quick steps → Manage quick steps**). Make one per folder you file into often, using the action *Move to folder*. You can add *Categorize* to the same step and give it a shortcut like Ctrl+Shift+1.

(Moving from the add-in itself would need a Microsoft app registration that IT may have to approve. We skipped that for now.)

## Files

| File | What it is |
|---|---|
| `manifest.xml` | Tells Outlook where the add-in lives. This is the file you install. |
| `taskpane.html / .js / .css` | The panel itself. |
| `assets/` | Icons. |
| `set-host.ps1` | Only needed if the add-in moves to a different web address: `.\set-host.ps1 -BaseUrl https://<new-address>` updates every URL in `manifest.xml`. |
