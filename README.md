# YACardEmu GUI

A lightweight graphical frontend for YACardEmu that simplifies card management, image handling, and card selection.

This tool provides an interface for working with `card.bin` files, assigning clean template images, and inserting cards without manually editing files.

---

## Features

* Automatic detection of new `card.bin` files
* Prompt to rename new cards and assign template images
* Thumbnail grid for selecting clean card templates
* Preview of current card image
* Next / Previous navigation
* One-click card insertion via YACardEmu API
* Card ↔ template linking system
* Reset card PNGs back to clean templates
* Persistent GUI settings (window size/position)
* Built-in card manager window

---

## Download

Download the latest release here:
➡️ https://github.com/YOUR_USERNAME/YOUR_REPO/releases/latest

---

## Requirements

* YACardEmu
* No Python required when using the prebuilt release

---

## Installation

1. Download and extract the latest release
2. Place the files inside your YACardEmu folder

Example:

```text id="layout1"
YACardEmu/
│
├── YACardEmu.exe
├── config.ini
├── YACardEmu_Gui.exe
├── data/
│   ├── yacardemu_gui.ini
│   ├── card_image_links.json
│   ├── card_images/
```

3. Run:

```text id="run1"
YACardEmu_Gui.exe
```

---

## Running from Source (Optional)

```text id="run2"
YACardEmu_Gui.pyw
```

Requires Python 3.x.

---

## How It Works

* Automatically launches YACardEmu
* Connects to local API (`127.0.0.1`)
* Reads `cards` folder from `config.ini`
* Detects new `card.bin` files
* Prompts for rename + template assignment
* Handles card insertion via API

---

## Screenshots

*(Add screenshots here)*

* Main window
* Template selection grid
* Card manager window

---

## Card Templates

Place clean PNG templates here:

```text id="templates"
/data/card_images/
```

* Filenames can be anything
* Used to overwrite card PNGs before insertion
* Prevents text stacking from repeated saves

---

## Data Files

```text id="datafiles"
/data
  yacardemu_gui.ini       # GUI settings
  card_image_links.json   # Card ↔ template mappings
  /card_images            # Template PNGs
```

---

## Managing Cards

Use **Manage Cards** to:

* Rename cards
* Link/unlink templates
* Reset PNGs from templates
* Fix new `card.bin` files
* Open template folder

---

## Building the Executable (Optional)

Prebuilt executables are provided in Releases.

### Tool

Use **PyInstaller**

### Install

```bash id="build1"
pip install pyinstaller
```

### Build

```bash id="build2"
pyinstaller --noconsole --onefile YACardEmu_Gui.pyw
```

### Output

```text id="build3"
/dist/YACardEmu_Gui.exe
```

### Notes

* No Python required for compiled version
* Keep `/data` folder next to `.exe`
* Must be in same folder as `YACardEmu.exe`

---

## Troubleshooting

### YACardEmu does not start

* Ensure `YACardEmu.exe` is in the same folder as the GUI
* Check that antivirus is not blocking it

---

### “Startup failed” error

* Verify `config.ini` exists
* Ensure `apiport` in `config.ini` is valid (default: 8080)
* Make sure the port is not already in use

---

### No cards detected

* Confirm `basepath` in `config.ini` points to the correct `cards` folder
* Ensure `.bin` files exist in that folder

---

### Images not showing

* Confirm `.png` exists alongside `.bin`
* Check that images are not corrupted

---

### Template selection is empty

* Add `.png` files to:

```
/data/card_images/
```

---

### New card prompt not appearing

* Ensure a file named `card.bin` is created in the cards folder
* The app checks for changes every ~1.5 seconds

---

### Insert button not working

* Ensure YACardEmu is running
* Verify API is reachable on `127.0.0.1:<port>`

---

## Changelog

### v1.0.0

* Initial release
* Card detection and renaming
* Template system
* Card insertion via API
* Card manager window

---

## License

This project is licensed under the GNU General Public License v2.0 (GPL-2.0).

It is designed to work alongside YACardEmu, which is also licensed under GPL-2.0.

See the LICENSE file for full details.

---

## Credits

YACardEmu:
https://github.com/GXTX/YACardEmu

All credit for the core card emulation functionality belongs to the original authors.

This project is an independent GUI frontend.

---

## Disclaimer

This project is provided as-is with no warranty.
Use at your own risk.
