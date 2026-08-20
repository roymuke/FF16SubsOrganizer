# FFXVI Subtitle Organizer
Script made for easier subtitle retranslation for FFXVI. Supports base and demo game versions, along with both DLCs.

> [!WARNING]
> Tested by repacking manually, modded subtitles have not been tested with [Reloaded-II Mod Manager](https://github.com/Reloaded-Project/Reloaded-II).
# Requirements
* [FF16Tools](https://github.com/Nenkai/FF16Tools)
* [FF16Converter 1.4](https://github.com/KillzXGaming/FF16Converter) (optional)
* Python (version used: 3.10.6)
	* pip (optional)
	* openpyxl
* Microsoft Excel (version used: Excel 2010)
# Usage
After you extract the contents of `0007.xx.pac` (`xx` being your selected language) and `0007.ja.pac` with `FF16Tools`, convert the files with `FF16SubsOrganizer`, generate your `xlsx`, edit it and modify the `xml` files, convert back and your are done. Check the [Wiki](https://github.com/roymuke/FF16SubsOrganizer/wiki) for a more detailed [step-by-step](https://github.com/roymuke/FF16SubsOrganizer/wiki/How-to-make-a-retranslation).

If you are working with the demo version, use `0001.xx.pac` and `0001.ja.pac`.

For DLC1 use `2002.xx.pac` and `2002.ja.pac`; for DLC2 use `3002.xx.pac` and `3002.ja.pac`.

> [!WARNING]
> This script doesn't have an *UI*, works only via command line.
## Commands
To convert `pzd` to `xml`, or `xml` to `pzd` in batch, optional command to move those files into another directory:
```shell
python FF16SubsOrganizer.py convert-batch (-c "<drive>:\path\to\FF16Converter.exe" | -b) -f "<drive>:\path\to\folder\0007.en\nxd\text" (--pzd | --xml) [-m "<drive>:\path\to\moving\folder"] [--verbose]
```
* `-c`: "FF16Converter.exe" directory path.
* `-b`: uses built-in converter.
* `-f`: Path to language folder.
* `--pzd`: Extension for files to convert, PZD to XML.
* `--xml`: Extension for files to convert, XML to PZD.
* `-m` (optional): Folder path to move newly generated `.pzd` or `.xml` files.
* `--verbose` (optional): show detailed output messages.
---
To extract `xml` dialogue and export to `xlsx` (excel):
```shell
python FF16SubsOrganizer.py to-xlsx -l "<drive>:\path\to\folder\0007.en" -j "<drive>:\path\to\folder\0007.ja" [-o "<drive>:\custom\path\to\file.xlsx"] [--verbose]
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
python FF16SubsOrganizer.py edit-xml -f "<drive>:\path\to\file.xlsx" -col I2 -l "<drive>:\path\to\folder\0007.en" [--verbose]
```
* `-f`: XLSX file directory.
* `-col`: column (and row) where new translation is located in the `xlsx` file. Recommended `I2`.
* `-l`: language folder directory to be translated.
* `--verbose` (optional): show detailed output messages.
---
Move files by extension to another directory:
```shell
python FF16SubsOrganizer.py move-batch -f "<drive>:\path\to\folder\0007.en.XML" (--pzd | --xml) -m "<drive>:\path\to\moving\folder" [--verbose]
```
* `-f`: Path to folder.
* `--pzd`: extension to move XML files.
* `--xml`: extension to move PZD files.
* `-m`: Destination folder path to move files.
* `--verbose` (optional): show detailed output messages.
# Thanks
* [KillzXGaming](https://github.com/KillzXGaming) for the original `.pzd` converter.
# Feedback
Did you use my script? Feel free to open an [issue ticket](https://github.com/roymuke/FF16SubsOrganizer/issues) in case you encountered any bug.

You can send me an email at [contact@roysu.cl](mailto:contact@roysu.cl) with your comments or find me in the "Final Fantasy XVI Modding" Discord server as *Roysu* (Roysu#7893).

Your feedback would be greatly appreciated!