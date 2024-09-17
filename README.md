# vATISLoad

_by Josh Glottmann_

**Version 1.2.8** - 09/15/2024

Fetches and loads AIRNC D-ATISs for use with [vATIS](https://vatis.clowd.io/) by [Justin Shannon](https://github.com/JustinShannon)

__[Download v1.2.8](https://github.com/glott/vATISLoad/releases/latest/download/vATISLoad.py)__ 

### Usage

1) Download and install the latest version of [Python](https://www.python.org/downloads/).
2) Launch `vATISLoad.py` and follow the prompts provided to create a configuration file and use the software. 
3) The facility name must be the same (or partially unique) to an existing vATIS facility.
4) __Do not move your mouse/type__ while ATIS data is being populated.

### Valid *Airports* Examples

- `SFO, OAK, SJC, SMF, RNO`

- `MIA/D, MIA/A, FLL`

- `OAK`

### Configuration File

- [Sample configuration file](https://github.com/glott/vATISLoad/blob/main/vATISLoadConfig.json)

- `vATISLoadConfig.json` is saved in either `%localappdata%\vATIS\` or `%localappdata%\vATIS-4.0\` 

- To delete profiles, remove the facility's section directly from the configuration file.