ToggleVM
ToggleVM is an interactive PowerShell script that lets you manage your Hyper-V virtual machine and its Parsec connection. It lets you start, stop, restart or force-kill your VM, launch or toggle Parsec, and manage VM checkpoints with live search and overwrite functionality. ToggleVM works best with the Easy-GPU-PV setup available at https://github.com/jamesstringerparsec/Easy-GPU-PV.

Features
VM Management:

Start, shut down (gracefully or forcefully), restart or kill your Hyper‑V VM.

Parsec Control:

Launch or toggle Parsec connectivity with configurable automatic closing behavior.

Checkpoint Management:

Create, apply, and overwrite checkpoints.

Live in‑menu search for checkpoints.

Settings:

Configure Parsec close policy (always close, never close, or prompt).

Toggle safety confirmations for destructive actions.

Interactive Menu:

Use arrow keys, hotkeys (n, o, s, b, etc.), and live search to quickly find and manage checkpoints.

Prerequisites
Windows with Hyper-V enabled.

Parsec installed.

PowerShell 5.1 (or later) installed.

Best used within the Easy-GPU-PV environment:
https://github.com/jamesstringerparsec/Easy-GPU-PV

Installation
Clone or download this repository.

Place the ToggleVM.ps1 script in your desired folder.

Open a PowerShell prompt with administrative privileges.

Run the script. On the first run it will generate a configuration file (GPUPV.config) and prompt for your Parsec peer ID if missing.

Usage
When you run the script, you’ll see an interactive menu with options to control the VM, manage Parsec, and handle checkpoints. You can use the following keys within the menu:

1, 2, etc.: Execute the corresponding main menu actions.

Arrow keys: Navigate through the menu.

n: Create a new checkpoint.

o: Overwrite the selected checkpoint (except for the “Factory Reset” checkpoint).

s: Toggle live search mode. When active, type to filter checkpoints instantly.

b: Return to the main menu.

Enter: Open the actions menu for the selected checkpoint.

You can also run the script with command-line switches (e.g. -StartVM, -ShutdownVM, etc.) to automate tasks without entering the menu.

Configuration
The script uses a configuration file named GPUPV.config that stores essential values such as:

VM name

Parsec executable path

Parsec close policy (alwaysClose, neverClose, or prompt)

Safety confirmation flag

Parsec peer ID (if not found, you will be prompted to enter it)

Feel free to modify these values or use the in-menu settings to update them.

Contributing
Contributions, feedback, and feature suggestions are welcome. Feel free to open issues or submit pull requests.

License
This project is released under [insert your license here].

Support
For support or questions, please open an issue on this repository.
