# vATISLoad

_by Josh Glottmann_

**Version 1.4.0** - 03/30/2025

Fetches and loads AIRNC D-ATISs for use with [vATIS](https://vatis.app/) by [Justin Shannon](https://github.com/JustinShannon)

__[Download v1.4.0](https://github.com/glott/vATISLoad/releases/latest/download/vATISLoad.pyw)__ 

### Usage

1) Download and install the latest version of [Python](https://www.python.org/downloads/).
2) Launch `vATISLoad.pyw`.
3) Select a profile in vATIS.
4) vATISLoad will automatically connect airports with a D-ATIS on initial load.
5) vATISLoad will automatically refresh D-ATIS information every minute until vATIS is shutdown.

\* _Note: the first time you run vATISLoad, it may take a few minutes to download the required packages to make the script run properly._

###  Replacements

- Replacements are defined in the [vATISLoadConfig.json](https://github.com/glott/vATISLoad/blob/main/vATISLoadConfig.json) file. 

- Replacements allow text to be stripped/modified from a D-ATIS to make it more suitable to regular usage.

- Each listed text will be replaced by the defined replacements. 

- Regex replacements are supported by including `%r` at the beginning of the replacement section (`%r` will be stripped out in the actual replacement). See `KMIA` or `KFLL` for valid regex replacement examples. 

- Replacements occur before any contractions are replaced in the text. 

- Submit a pull request to add additional replacement data.
