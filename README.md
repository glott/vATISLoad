# vATISLoad

_by Josh Glottmann_

**Version 1.5.0** - 10/23/2026

Automatically selects the correct profile and connects D-ATIS stations in [vATIS](https://vatis.app/) by [Justin Shannon](https://github.com/JustinShannon). vATIS now fetches and applies D-ATIS data natively, so vATISLoad no longer uplinks any text itself; it just gets the right stations online with the `D-ATIS` preset selected.

__[Download v1.5.0](https://github.com/glott/vATISLoad/releases/latest/download/vATISLoad.pyw)__

### Installation

1) Download and install the latest version of [Python](https://www.python.org/downloads/). Select `Add python.exe to PATH` when installing Python.
2) Launch `vATISLoad.pyw` to download all required libraries.
3) Add a `D-ATIS` preset to each station that you would like connected automatically. Stations without this preset are ignored.

\* _Note: the first time you run vATISLoad, it may take a few minutes to download the required libraries to make the script run properly._

### Usage
1) Open CRC and connect to the network.
2) Launch `vATISLoad.pyw`.
3) vATISLoad will:
    - Open vATIS if it is not already running.
    - Select the vATIS profile matching the ARTCC of your active CRC profile (ZOA, ZMA, and ZDC by default; set `AUTO_SELECT_FACILITY = True` in the script to enable this for any ARTCC). If no match is found, select a profile in vATIS manually.
    - Select the `D-ATIS` preset on every station that has one.
    - Connect those stations one at a time, waiting for each to report `Connected` before starting the next, up to a maximum of 4 simultaneous connections.
    - If you are logged into a TWR, GND, DEL, or RMP position in CRC, only that airport's stations are connected.
4) vATISLoad exits once the ATISes are connected. vATIS keeps the D-ATIS data refreshed on its own.

\* _Note: if vATIS is updating, vATISLoad may fail to start vATIS properly. Launch vATIS first, wait for it to update, then launch vATISLoad again._

### Replacements

- vATISLoad no longer modifies D-ATIS text. Text replacements are now handled by vATIS itself through each station's `datisTextReplacements` setting in the vATIS profile.
- The previous [vATISLoadConfig.json](https://github.com/glott/vATISLoad/blob/main/vATISLoadConfig.json) replacement list has been migrated into the ZOA and ZMA profiles. The file is kept for reference only; the script does not read it.
- To add or change a replacement, edit the station's `datisTextReplacements` in the vATIS profile (`pattern`, `replacement`, and `isRegex`) and submit a pull request to the facility's profile repository.

### Troubleshooting

- Is the script doing nothing the first time you run it?
  - It is likely installing required libraries to run. You will see a popup the first time this happens. If you don't see a popup, that means the required libraries are installed.
  - Alternatively, run [vATISLoad_library_installer.py](https://github.com/glott/vATISLoad/raw/refs/heads/main/archive/vATISLoad_library_installer.py). This will attempt to install required libraries. After running this script, attempt to run `vATISLoad.pyw` again.

- Is vATIS being opened automatically by the script?
  - If yes, then that's a good sign!

- Was the correct profile selected?
  - Auto-selection requires CRC to be open with a profile loaded. If the wrong profile (or none) was selected, pick it manually in vATIS and run vATISLoad again.

- Does each station you expect to connect have a preset named `D-ATIS`?
  - Only stations with a `D-ATIS` preset are connected.

- Did your ATISes connect?
  - If yes, then you're all set! If no, make sure you have an active connection in CRC.
  - vATISLoad will not connect more than 4 ATISes. If 4 are already connected, it will not connect any more.

- For further troubleshooting, try the following steps:

  1. Quit vATIS
  2. Rename `vATISLoad.pyw` to `vATISLoad.py` (dropping the `w`). If you can’t see the file extension, open the `View` menu in File Explorer then tick `File name extensions` in the `Show/hide` section.
  3. Open `Command Prompt` and drag the `vATISLoad.py` file into it, press enter to run it. If vATIS opens automatically, select whichever profile has your D-ATIS airports. The script prints each station as it connects (or fails to).
  4. Send Josh a screenshot of `Command Prompt` after following these steps.