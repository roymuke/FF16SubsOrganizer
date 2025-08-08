# FFXVI Subtitle Organizer
Script made for easier subtitle retranslation for FFXVI.

> [!WARNING]
> * This script was made with the DEMO version of FFXVI, can't confirm if it works properly on the base version. If you try it, use `0007.xx`.
> * Tested by repacking manually, hasn't been tested with [Reloaded-II Mod Manager](https://github.com/Reloaded-Project/Reloaded-II).
# Requirements
* [FF16Tools](https://github.com/Nenkai/FF16Tools)
* [FF16Converter 1.4](https://github.com/KillzXGaming/FF16Converter)
* Python (version used: 3.10.6)
	* pip (optional)
	* openpyxl
* Microsoft Excel (version used Excel 2010)
# Usage
After you extracted the content of `0001.xx.pac` and `0001.ja.pac` with `FF16Tools`, convert the `.pzd` files to `.xml` with `FF16Converter` (found, by deafult, in path `0001.xx.pac\nxd\text`).
## Commands
To extract `xml` dialogue and export to `xlsx` (excel):
```shell
python FF16SubsOrganizer.py to-xlsx -l "<drive>:\path\to\folder\0001.en" -j "<drive>:\path\to\folder\0001.ja" [-o "<drive>:\custom\path\to\file.xlsx"]
```
* `-l`: language folder directory for translation.
* `-j`: japanese folder directory.
* `-o` (optional): output directory, by default it's on same directory as the script.

> [!IMPORTANT]
> When editing xlsx...
> * be mindful of `<br>`, always add a newline after one, I haven't checked what happens if you don't add one.
> * after pasting dialogue with newlines, select the table and separate cells, some rows may get hidden after pasting and this should fix it.

---
To convert `xlsx` back to `xml`:
```shell
python FF16SubsOrganizer.py edit-xml -f "file.xlsx" -col I2 -l "<drive>:\path\to\folder\0001.en"
```
* `-f`: XLSX file directory.
* `-col`: column (and row) where starts user retranslation, title column doesn't count. Recommended `I2`.
* `-l`: language folder directory, from where user wants to translate.

> [!NOTE]
> #### Directory Organization
> Depending on the way you organized your extracted XML files, you will need to use an specific path, the script looks for the **direct parent folder** of `bevent`, `bossbattle`, etc. If you extracted the XML on the same folder as the PZD ones, your path would be `C:\0001.en\nxd\text`; If you put the XML files on a separated folder, let's say: `C:\custom 0001 english XML\`, where inside are the folders `bevent`, `bossbattle`, etc. Then you should use `C:\custom 0001 english XML` as path.

---
To extract `xml` dialogue and save it on a `html` file (this is just a leftover funcionality):
```shell
python FF16SubsOrganizer.py to-html -l "<drive>:\path\to\folder\0001.en" -j "<drive>:\path\to\folder\0001.ja" [-o "<drive>:\custom\path\to\file.html"]
```
* `-l`: language folder directory of interest.
* `-j`: japanese folder directory.
* `-o` (optional): output directory, by default it's on same directory as the script.
