# Silent Command Autorun Utility

It is surprisingly difficult to run commands or scripts at startup/login
on a Windows system without an unpleasant command-line window popping up.
The simplest way to achieve a silent launch is by executing a Visual Basic
Script on boot, with the necessary commands enclosed inside it.

This small utility (itself a .vbs file) guides the user through the process of generating and
setting up such a script, with easy to understand, interactive GUI prompts.
No programming knowledge or experience is required.

## Screenshots

![screenshot_1](/assets/images/prompt1.png "Start confirm")

![screenshot_2](/assets/images/prompt2.png "Number of commands")

![screenshot_3](/assets/images/prompt3.png "Enter command")

![screenshot_4](/assets/images/prompt4.png "Copy prompt")

## Instructions

Run silent_cmd_autorun.vbs, follow the prompts, and step by step input the
required data to generate the needed script.

## Warning

The utility offers the option to copy the completed script into the current
user's startup directory. Some anti-malware software may label this as
"suspicious activity".

## Other

Tested on Windows 11, with Windows Script Host version 5.812.

**[Contact](mailto:lcs_it@proton.me)**

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
