# desktop-file-organizer
### your Desktop is not a filing system. it is a cry for help.

> Two Python scripts. One suspiciously messy Desktop. A tiny attempt to restore order to the filesystem before it achieves sentience.

This is a small, no-dependencies Python toolkit for people who collect files like browser tabs, download things with names like `document_final_FINAL2.pdf`, and occasionally remember that folders exist.

It inventories the chaos first, then helps sort the loot into sensible Portuguese-named folders. Think **school-office bureaucracy meets terminal goblin**: methodical enough to make a report, unhinged enough to need one.

## ✨ The lore

The internet has enough overengineered productivity apps. This project is two scripts and a dream:

- **`inventario_arquivos.py`** — the forensic investigator. Recursively lists files and folders so you can see what kind of digital ecosystem you're living in.
- **`organizar_area_trabalho.py`** — the tiny janitor with a clipboard. Previews file moves, sorts files into categories, and cleans up directories that become empty.

No cloud account. No subscription. No AI pretending that renaming `thing2_new_revised_REAL.pdf` is a spiritual journey. Just Python and your filesystem.

## 🧰 What it does

- **Recursive inventory:** makes a readable report of files and folders, including nested directories.
- **Rule-based sorting:** uses filenames, paths, extensions, and recognizable keywords to choose destinations. It is practical heuristics, not mind-reading.
- **Portuguese folder names:** categories are descriptive and designed to be easy to find, including school administration, Censo Escolar, attendance, math/study materials, programming projects, PDFs, spreadsheets, images, archives, and assorted digital debris.
- **Protected zones:** the organizer is designed to leave folders whose names contain `lagora`, `censo`, or `sistema presença` (including common accent variations) alone.
- **`shutdown.bat` gets plot armor:** files with that name are protected, along with the folders needed to keep them where they are.
- **Preview before action:** it shows the proposed moves and requires the exact confirmation `ORGANIZAR`. Anything else cancels. No surprise speedrun.
- **No silent overwrites:** when a destination already has a file with the same name, the organizer uses a different name instead of replacing it.
- **CSV move log:** writes `registro_organizacao.csv` so you can inspect what moved where.
- **Empty-folder cleanup:** removes directories only when they are empty after the moves. Files are not supposed to be deleted as part of this cleanup.

## 🚀 Getting started

### Requirements

- Windows is the main target.
- Python 3.10+ is recommended.
- No third-party packages. The standard library is doing the heavy lifting, as nature intended.

Check that Python is available:

```powershell
python --version
```

### 1. Put the scripts somewhere sensible

Download or clone this repository, then place the scripts in a folder you can find again. You may run them from your Desktop or a separate tools folder.

```powershell
git clone https://github.com/YOUR-USERNAME/desktop-file-organizer.git
cd desktop-file-organizer
```

Replace `YOUR-USERNAME` with your GitHub username, and make sure the inventory script's actual filename matches the command below.

### 2. Take inventory before touching anything

```powershell
python inventario_arquivos.py
```

Read the script's prompts and output path; the exact output filename depends on the inventory script version you're using. Treat the inventory like a pre-flight checklist: inspect the terrain before launching the file-moving machinery.

### 3. Preview the reorganization

```powershell
python organizar_area_trabalho.py
```

The organizer attempts to detect your Desktop, including common localized and OneDrive paths. It should show a preview of the planned moves before changing anything.

Read it. Actually read it. This is the part where you get to be smarter than a raccoon with administrator privileges.

If the plan looks right, type:

```text
ORGANIZAR
```

Anything else cancels the operation.

Afterward, inspect `registro_organizacao.csv` to see the recorded moves.

## 🛡️ Protected items: plot armor edition

The organizer is intended to leave these items untouched:

| Item | Protection |
|---|---|
| Any folder whose name contains `lagora` | The folder and its contents are skipped |
| Any folder whose name contains `censo` | The folder and its contents are skipped |
| Any folder whose name contains `sistema presença` | The folder and its contents are skipped, including common accent variations |
| Any file named `shutdown.bat` | The file is skipped, and its containing folder is preserved as needed |

The inventory script may still list protected items because its job is to report what exists. The protections above apply to the organizer.

**Important:** protection is based on names and paths, not on understanding what an application is. If a folder is mission-critical, check that its name matches the protection rules and verify the preview.

## 🗂️ How the sorting brain works

The organizer uses keyword and extension rules to make educated guesses. A file about school attendance may land in a school-related folder; a `.pdf` with no stronger clue may land in a general PDF folder.

This is not semantic AI. It cannot reliably understand every cryptic filename, weird extension, or file whose entire identity is `new_new2`. Check the preview and manually fix any questionable placements.

## ⚠️ Safety briefing from the goblin union

Moving files can break shortcuts, scripts, project references, or apps that expect files to stay in their original locations.

Before running the organizer:

1. **Back up important files.** This is not optional wisdom; it is the ancient law of computers.
2. Generate and review the inventory.
3. Run the organizer and inspect the preview.
4. Cancel if anything looks wrong.
5. After organizing, verify important projects and inspect the CSV log.

A CSV log is an audit trail, **not an automatic undo button**. The script is intended to move files, not delete their contents, but don't treat any bulk file operation as risk-free.

## 🧪 Suggested workflow

```text
inventory → inspect the chaos → preview the moves → confirm → audit the log
```

In other words: measure twice, yeet once.

## 🛠️ Customization

Want different folder categories? Edit the organizer's keyword and extension rules. Want different protected folders? Update its protection checks and test the preview before using it on real files.

If you improve the classification rules, make them readable. Future-you is a separate person with no memory of why you added a mysterious condition at 2:17 a.m.

## 🤝 Contributions

Bug reports, clearer category rules, better filename handling, and safety improvements are welcome. Keep changes boring where data safety is concerned and silly where documentation is concerned.

## 📜 License

No license has been selected yet. Add a `LICENSE` file before presenting this repository as open source; without one, others do not automatically receive permission to reuse, modify, or redistribute the code.

---

*Made for the noble quest of turning Desktop entropy into folders. The files were always there. They were simply choosing violence.*
