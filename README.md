# vATISLoad

_by Josh Glottmann_

**Version 1.3.2** - 01/27/2025

Fetches and loads AIRNC D-ATISs for use with [vATIS](https://vatis.app/) by [Justin Shannon](https://github.com/JustinShannon)

__[Download v1.3.2](https://github.com/glott/vATISLoad/releases/latest/download/vATISLoad.pyw)__ 

### Usage

1) Download and install the latest version of [Python](https://www.python.org/downloads/).
2) Ensure your vATIS profile name has the facility name in parentheses (e.g. `Oakland ARTCC (ZOA)`). This is essential to automatically determining the vATIS profile to select. Square brackets (`[]`) are also accepted to surround the facility name. 
3) Ensure each station in the vATIS profile has its first preset named `D-ATIS`.
4) Launch `vATISLoad.pyw` after launching CRC.
5) __Do not move your mouse/type__ while ATIS data is being populated. Your mouse may automatically be frozen for your convenience. *
6) Connect ATISes at your discretion.

\* _Note: the first time you run vATISLoad, it may take a few minutes to download the required packages to make the script run properly._

### Facility Patches

- Facility patches are defined in the [vATISLoadConfig.json](https://github.com/glott/vATISLoad/blob/main/vATISLoadConfig.json) file. 

- Patches allow people signed onto a lower level position to utilize the higher level D-ATIS vATIS profile.

- For example, a person signed onto a `ZOA`, `NCT`, `SFO`, `OAK`, `SJC`, `SMF`, or `RNO` position would be patched to the `Oakland ARTCC (ZOA)` vATIS profile.

- This does not work in the other direction - a `San Francisco ATCT (SFO)` vATIS profile would only be selected for a `SFO` controller. 

- Submit a pull request to add additional facility patching data.

###  Replacements

- Replacements are defined in the [vATISLoadConfig.json](https://github.com/glott/vATISLoad/blob/main/vATISLoadConfig.json) file. 

- Replacements allow text to be stripped/modified from a D-ATIS to make it more suitable to regular usage.

- Each listed text will be replaced by the defined replacements. 

- Regex replacements are supported by including `%r` at the beginning of the replacement section (`%r` will be stripped out in the actual replacement). See `KMIA` or `KFLL` for valid regex replacement examples. 

- Replacements occur before any contractions are replaced in the text. 

- Submit a pull request to add additional replacement data.
