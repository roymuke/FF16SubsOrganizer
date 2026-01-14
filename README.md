# FFXVI Subtitle Organizer
Script made for easier subtitle retranslation for FFXVI. Supports both base and demo game versions.

> [!WARNING]
> Tested by repacking manually, modded subtitles have not been tested with [Reloaded-II Mod Manager](https://github.com/Reloaded-Project/Reloaded-II).
# Requirements
* [FF16Tools](https://github.com/Nenkai/FF16Tools)
* [FF16Converter 1.4](https://github.com/KillzXGaming/FF16Converter)
* Python (version used: 3.10.6)
	* pip (optional)
	* openpyxl
* Microsoft Excel (version used: Excel 2010)
# Usage
After you extract the contents of `0007.xx.pac` (`xx` being your selected language) and `0007.ja.pac` with `FF16Tools`, convert the files with `FF16SubsOrganizer`, generate your `xlsx`, edit it and modify the `xml` files, convert back and your are done. Check the [Wiki](https://github.com/roymuke/FF16SubsOrganizer/wiki) for a more detailed [step-by-step](https://github.com/roymuke/FF16SubsOrganizer/wiki/How-to-make-a-retranslation).

If you are working with the demo version, use `0001.xx.pac` and `0001.ja.pac`.

> [!WARNING]
> This script doesn't have an *UI*, works only via command line.
## Commands
To convert `pzd` to `xml`, or `xml` to `pzd` in batch, optional command to move those files into another directory:
```shell
FF16SubsOrganizer.py convert-batch -c "<drive>:\path\to\FF16Converter.exe" -f "<drive>:\path\to\folder\0007.en\nxd\text" --pzd [-m "<drive>:\path\to\moving\folder"]
```
* `-c`: "FF16Converter.exe" directory path.
* `-f`: Path to language folder.
* `--pzd`: Extension to convert, i.e, PZD to XML. (has to be just one of these)
* `--xml`: Extension to convert, i.e, XML to PZD. (has to be just one of these)
* `-m` (optional): Folder path to move newly generated `.pzd` or `.xml` files.
* `--verbose` (optional): show detailed output messages.
---
To extract `xml` dialogue and export to `xlsx` (excel):
```shell
FF16SubsOrganizer.py to-xlsx -l "<drive>:\path\to\folder\0007.en" -j "<drive>:\path\to\folder\0007.ja" [-o "<drive>:\custom\path\to\file.xlsx"]
```
* `-l`: language folder directory for translation.
* `-j`: japanese folder directory.
* `-o` (optional): output directory, by default it's on same directory as the script.
* `--verbose` (optional): show detailed output messages.

> [!IMPORTANT]
> When editing the `xlsx` file, be mindful of `<br>`, always add a newline after one, I haven't checked what happens if you don't add one.

---
To convert `xlsx` back to `xml`:
```shell
FF16SubsOrganizer.py edit-xml -f "<drive>:\path\to\file.xlsx" -col I2 -l "<drive>:\path\to\folder\0007.en"
```
* `-f`: XLSX file directory.
* `-col`: column (and row) where starts user retranslation, title column doesn't count. Recommended `I2`.
* `-l`: language folder directory, from where user wants to translate.
* `--verbose` (optional): show detailed output messages.

> [!NOTE]
> #### Directory Organization
> Depending on the way you organized your extracted XML files, you will need to use an specific path, the script looks for the **direct parent folder** of `bevent`, `bossbattle`, etc. If you extracted the XML on the same folder as the PZD ones, your path would be `C:\0007.en\nxd\text`; If you put the XML files on a separated folder, let's say: `C:\custom 0007 english XML\`, where inside are the folders `bevent`, `bossbattle`, etc. Then you should use `C:\custom 0007 english XML` as path.

---
Move files by extension to another directory:
```shell
FF16SubsOrganizer.py move-batch -f "<drive>:\path\to\folder\0007.en.XML" --pzd -m "<drive>:\path\to\moving\folder"
```
* `-f`: Path to folder.
* `--pzd`: PZD extension files to move. (has to be just one of these)
* `--xml`: XML extension files to move. (has to be just one of these)
* `-m`: Destination folder path to move files.
* `--verbose` (optional): show detailed output messages.
# Feedback
Did you use my script? Feel free to open an [issue ticket](https://github.com/roymuke/FF16SubsOrganizer/issues) in case you encountered any bug.

You can send me an email at [contact@roysu.cl](mailto:contact@roysu.cl) with your comments or find me in the "Final Fantasy XVI Modding" Discord server as *Roysu* (Roysu#7893).

Your feedback would be greatly appreciated!