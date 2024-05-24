# KolMafia Launcher

KolMafia Launcher is a launcher/updater for KoLMafia launcher.

## Support

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/G2G6YI7DO)

## Features

* Automatically determines the system-defined launcher for GUI-based JAR files, and uses that to launch KoLMafia
* Checks for updated KoLMafia builds on launch and, if found, automatically downloads them
* Allows the user to have a static shortcut for launching KoLMafia, even though the target name changes on a regular basis
* Self-updating capability with user opt-out option
* Optional installation of scheduled task that runs during rollover to potentially speed launch
* Digitally signed binary and installer from Microsoft-trusted CA

## Installation

### Binary

1. Download and run the KoLMafia-Launcher executable in the root of the repository or attached to the latest release

## Usage

* On first use, KoLMafia Launcher will request the location of your current KoLMafia installation
* Thereafter, KoLMafia will use that location as its baseline for version checking, and download location for new versions

### Command-line switches

`--noLaunch`
* Only runs update check, does not launch KolMafia

`--killOnUpdate`
* Checks for other instances of the system-defined launcher for GUI-based JAR files, and terminates them. Useful with the -Silent flag, as updating the JAR while it is running is not recommended

`--Silent`
* Prevents any error messages from appearing. Useful in conjunction with scheduled tasks

`--Verbose`
* Adds a number of logging statements to the console. Useful for assisting with troubleshooting.

## Legal Notice

icon.ico is a derivative artwork, based on IP owned in whole by Asymmetric Publications, LLC. This derivative has been created with permission from Asymmetric Publications, LLC. The author of KoLMafia-Launcher does not have the capacity to permit sub-derivative works, or to permit use of such work in other projects without authors of those projects obtaining express written consent of Asymmetric Publications, LLC.

### Translation

tptb were good enough to let me whip up an icon that incorporates their IP for use with this project. That doesn't mean YOU can use the icon for your project (since it contains THEIR IP), or make your own art, based on this art (and honestly, why would you want to). If you want to do something like that, just email them and ask yourself!

## Requirements

* .NET 7.0 for base installer
* None for standalone installer
