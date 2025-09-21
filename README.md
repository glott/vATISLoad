# vATISLoad

_by Josh Glottmann_

**Version 1.4.6** - 09/20/2025

Fetches and loads AIRNC D-ATISs for use with [vATIS](https://vatis.app/) by [Justin Shannon](https://github.com/JustinShannon). D-ATIS data is automatically uplinked to vATIS after selecting a profile and gets refreshed every 15 minutes. 

__[Download v1.4.6](https://github.com/glott/vATISLoad/releases/latest/download/vATISLoad.pyw)__ 

### Installation

1) Download and install the latest version of [Python](https://www.python.org/downloads/). Select `Add python.exe to PATH` when installing Python.
2) Launch `vATISLoad.pyw` to download all required libraries.
3) Add a `D-ATIS` preset to each airport that you would like D-ATIS data uplinked to. If this preset does not exist, no information will be uplinked. 

\* _Note: the first time you run vATISLoad, it may take a few minutes to download the required libraries to make the script run properly._

### Usage
1) Open CRC and connect to the network.
2) Launch `vATISLoad.pyw`.
3) Select a profile in vATIS.
4) vATISLoad will attempt to uplink D-ATIS data and connect ATISes for you.
    - D-ATIS data will be refreshed every 5 minutes until vATIS is shutdown.
    - Within the first 5 minutes of launching vATISLoad, any text deleted from the `AIRPORT CONDITIONS` or `NOTAMS` sections of a station will be persistently deleted until restarting vATISLoad. Practically, this means you can delete "junk" text and it will not be included in any future D-ATIS updates. 

\* _Note: if vATIS is updating, vATISLoad may fail to start vATIS properly. Launch vATIS first, wait for it to update, then launch vATISLoad again._

###  Replacements

- Replacements are defined in the [vATISLoadConfig.json](https://github.com/glott/vATISLoad/blob/main/vATISLoadConfig.json) file. 

- Replacements allow text to be stripped/modified from a D-ATIS to make it more suitable to regular usage.

- Each listed text will be replaced by the defined replacements. 

- Regex replacements are supported by including `%r` at the beginning of the replacement section (`%r` will be stripped out in the actual replacement). See `KMIA` or `KFLL` for valid regex replacement examples. 

- Replacements occur before any contractions are replaced in the text. 

- Submit a pull request to add additional replacement data.

### Troubleshooting

- Is the script doing nothing the first time you run it?
  - It is likely installing required libraries to run. You will see a popup the first time this happens. If you don't see a popup, that means the required libraries are installed.
  - Alternatively, run [vATISLoad_library_installer.py](https://github.com/glott/vATISLoad/raw/refs/heads/main/archive/vATISLoad_library_installer.py). This will attempt to install required libraries. After running this script, attempt to run `vATISLoad.pyw` again.
 
- Is vATIS being opened automatically by the script?
  - If yes, then that's a good sign!
 
- If vATIS was opened automatically, did you select a profile?
  - You must select a profile for D-ATIS information to be uplinked.
 
- If you selected a profile, does each D-ATIS airport have a preset named `D-ATIS`?
  - Each airport much have a `D-ATIS` preset where information is uplinked to.

- If D-ATIS information was uplinked, did your ATISes connect?
  - If yes, then you're all set! If no, make sure you have an active connection in CRC. 

- For futher troubleshooting, try the following steps:

  1. Quit vATIS
  2. Rename `vATISLoad.pyw` to `vATISLoad.py` (dropping the `w`). If you can’t see the file extension, open the `View` menu in File Explorer then tick `File name extensions` in the `Show/hide` section.
  3. Open `Command Prompt` and drag the `vATISLoad.py` file into it, press enter to run it. If vATIS opens automatically, select whichever profile has your D-ATIS airports.
  4. Send Josh a screenshot of `Command Prompt` after following these steps. 
