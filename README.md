ToggleVM

A simple, fully offline, powershell script to control Hyper-V GPU-PV VMs and Parsec connections easily.

Best used together with Enhanced-GPU-PV: https://github.com/timminator/Enhanced-GPU-PV

-------------------------------------------------------------------------------

Features:

- Start, shutdown, restart, or force-kill a Hyper-V VM
- Connect to VM via Parsec automatically using a peer ID
- Control Parsec behavior: always close, never close, or prompt on shutdown
- Manage VM checkpoints: create, delete, rename, apply, overwrite
- Live search through checkpoints
- Factory reset the VM to a checkpoint
- Safety confirmation system for dangerous actions
- Config file system that saves settings automatically
- Works fully from a simple interactive menu
- Supports direct powershell flags if you don't want to use the menu

-------------------------------------------------------------------------------

Requirements:

- Windows with Hyper-V installed
- PowerShell 5.1 or later
- Parsec installed
- Administrative privileges (must run powershell as administrator)

-------------------------------------------------------------------------------

First Time Setup:

- When running for the first time, the script will generate a config file
- You will be prompted to input your Parsec peer ID if missing
- After that, everything is automatic

-------------------------------------------------------------------------------

Temporary Hyper-V Connection:

- There is an option to create a temporary Hyper-V connection that auto-closes after 25 seconds
- This forces the VM to properly register the external GPU display
- If you only connect with Parsec, sometimes it connects to the wrong display (Hyper-V's internal display)
- This would bypass the GPU completely and ruin performance
- The temporary connection ensures Parsec sees the external GPU display

Important tip: I recommend setting up a scheduled task inside the VM that switches to the external display on boot. This is what I do to make sure everything is always on the right display automatically.

-------------------------------------------------------------------------------

Command Line Flags:

- -StartVM : start the VM and connect via Parsec
- -ShutdownVM : gracefully shutdown the VM
- -KillVM : force kill the VM immediately
- -ConnectParsec : connect to the VM via Parsec
- -OpenMenu : forces the interactive menu to open

-------------------------------------------------------------------------------

Settings You Can Change:

- Set new Parsec peer ID
- Set Parsec close behavior (alwaysClose, neverClose, prompt)
- Toggle safety confirmation (on/off)
- Manage VM checkpoints (create, delete, rename, overwrite, factory reset)

-------------------------------------------------------------------------------

Notes:

- If Parsec is missing, wrong peer ID, or VM name wrong, you can change it easily inside the config file
- The script creates and uses "GPUPV.config" and "GPUPV.log" automatically
- Factory reset option only works if you created a checkpoint called "factory reset"

-------------------------------------------------------------------------------

Important:

Always run the script as administrator or Hyper-V commands and Parsec operations might fail.
