# FF16SubsOrganizer — C++ port

C++17 port of the Python script at the repo root. Same commands, same flags,
same behavior. Library equivalences:

| Python                   | C++                                |
| ------------------------ | ---------------------------------- |
| `xml.etree.ElementTree`  | [pugixml](https://pugixml.org/)    |
| `json`                   | [nlohmann/json](https://json.nlohmann.me/) |
| `argparse`               | [CLI11](https://github.com/CLIUtils/CLI11) |
| `openpyxl`               | [OpenXLSX](https://github.com/troldal/OpenXLSX) |
| `os`, `pathlib`, `shutil`| `std::filesystem`                  |
| `subprocess.run`         | `std::system` (with arg quoting)   |

All dependencies are fetched automatically via CMake `FetchContent`. No vcpkg /
Conan setup required.

## Build

Requires CMake >= 3.20 and a C++17 compiler (MSVC 2019+, GCC 9+, Clang 10+).

```sh
cd cpp
cmake -B build
cmake --build build --config Release
```

The resulting binary is `build/Release/FF16SubsOrganizer.exe` (MSVC) or
`build/FF16SubsOrganizer` (single-config generators).

`IDs.json` is loaded from the current working directory, so run the binary
from the repo root (where `IDs.json` lives) or copy `IDs.json` next to where
you invoke it.

## Usage

Identical to the Python script:

```sh
# Export subtitles to XLSX
FF16SubsOrganizer to-xlsx -l "C:/path/to/0007.en.XML" -j "C:/path/to/0007.ja/nxd/txt" -o file.xlsx

# Apply translations back to XML
FF16SubsOrganizer edit-xml -f file.xlsx -col I2 -l "C:/path/to/0007.en.XML"

# Batch convert pzd <-> xml
FF16SubsOrganizer convert-batch -c FF16Converter.exe -f "C:/path/to/folder" --pzd -m "C:/dest"

# Move files by extension
FF16SubsOrganizer move-batch -f "C:/src" --pzd -m "C:/dest"
```

## Notes on differences from the Python version

- The XLSX is generated with the same data layout (per-folder sheets, STATS
  sheet with progress formulas, hidden helper rows, column widths, freeze
  semantics). Advanced cell styling that `openpyxl` exposes trivially
  (alternating row fills, custom fonts on the STATS sheet) is omitted —
  OpenXLSX 0.4.x doesn't expose those APIs cleanly. Data and formulas are
  identical.
- ANSI colors require Windows 10 1607+ (auto-enabled via `SetConsoleMode`
  on startup).
- `convert-batch` chunks files (400 per call) for *all* folders, not just
  `defaultq` / `simpleq`, to stay safely under the Windows command-line
  length limit when calling `FF16Converter.exe`.
