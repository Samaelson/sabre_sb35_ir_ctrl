# sabre_sb35_ir_ctrl - Harman Kardon SABRE SB35 IRDROID Based Infrared Control System

A Windows-based infrared (IR) control system that allows you to control Harman Kardon SABRE SB35 sound system using simple text commands.

It uses a PowerShell client, a Python Windows Service backend, and a IRDROID USB IR Transceiver, communicating via Windows Named Pipes (JSON IPC).
---


# SYSTEM OVERVIEW

Batch Script (optional launcher)
        ↓
PowerShell Client (Script 1)
        ↓  Named Pipe (\\.\pipe\irdroid)
Python Windows Service (Script 2)
        ↓
USB IR Device (IRDroid USB Transceiver)
        ↓
Target Devices (Harman Kardon Sabre SB 35)
---

# FEATURES

## PowerShell Client:
- Command alias system (e.g. volup, hdmi1, bluetooth)
- Converts simple commands into IR actions
- Sends JSON requests via Named Pipe
- Waits for response (ok / fail / timeout)
- Supports multiple commands in one call
- Batch-friendly execution

## Python Windows Service:
- Runs as a Windows Service
- Auto-detects USB IR device (VID/PID)
- Serial communication with hardware
- IR command execution engine
- Multi-threaded processing (worker, pipe, watchdog)
- Automatic USB recovery (device reset via PowerShell)
- Metrics tracking (latency, success, failures)
- Request timeout handling
- Stale request cleanup (watchdog)

## Batch Script:
- Simple shortcut launcher
- Executes PowerShell commands hidden
---

# REQUIREMENTS

## Hardware:
- IRDroid USB IR Transceiver (or compatible device)

## Software:
- Windows 10 / 11
- Python 3.9+
- PowerShell 5+

## Python packages:
pip install pywin32 pyserial
---

# INSTALLATION

## 1. Clone repository:
```bash
git clone https://github.com/yourname/irdroid-control-system.git
cd irdroid-control-system
```

### 2. Install dependencies:
```bash
pip install pywin32 pyserial
```

## 3. Install Windows Service (run as Administrator):
```bash
python service.py install
python service.py start
```

## Optional auto-start:
```bash
python service.py --startup auto install
```

## 4. Connect hardware:
USB IR device must be connected.

Expected:
VID = 0x04D8
PID = 0xFD08


###5. Setup PowerShell client:
Example location:
```bash
C:\Scripts\irdroid.ps1
```
Test:
```bash
.\irdroid.ps1 power
.\irdroid.ps1 hdmi1 bluetooth
```

### 6. Optional batch launcher:
```bash
@echo off
powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass ^
-File "C:\Scripts\irdroid.ps1" -Command "bluetooth" "power"
```
---

# USAGE

## Basic commands:
```bash
.\irdroid.ps1 power
.\irdroid.ps1 hdmi1
.\irdroid.ps1 hdmi2
.\irdroid.ps1 volup
.\irdroid.ps1 voldown
.\irdroid.ps1 mute
```

## Multiple commands:
```bash
.\irdroid.ps1 bluetooth power hdmi1
```

---

# COMMAND MAPPING

power       -> on_off
toggle      -> on_off
hdmi1       -> hdmi_1
hdmi2       -> hdmi_2
hdmi3       -> hdmi_3
bluetooth   -> bluetooth
volup       -> vol_up
voldown     -> vol_down
mute        -> de_mute
bassup      -> bass_up
bassdown    -> bass_down
stereo      -> stereo
virtual     -> virtual

---

# COMMUNICATION PROTOCOL

Named Pipe:
\\.\pipe\irdroid

Request:
{
  "id": 1,
  "cmd": "bluetooth"
}

Response:
{
  "id": 1,
  "status": "ok"
}

Statuses:
- ok
- fail
- unknown
- timeout
---

# PYTHON SERVICE DETAILS

## Device Initialization:
- Detects USB device via VID/PID
- Validates firmware (v225)
- Validates mode (S01)
- Opens serial connection (115200 baud)

## IR Transmission:
- Sends start byte 0x03
- Sends data in 62-byte chunks
- Uses IR command mapping file

## Reliability:
- Automatic USB recovery (disable/enable device)
- Multi-threaded processing
- Watchdog cleanup
- Timeout handling (5 seconds)
- Safe locking system
- Graceful shutdown
---

# METRICS

## Tracks:
- ok (successful commands)
- fail (failed commands)
- unknown (invalid commands)
- timeout (no response)
- total commands
- latency per command

---

# LIMITATIONS

- No authentication (local only)
- Windows only
- Requires admin rights for recovery
- Depends on specific USB hardware (VID/PID)
- No encryption for IPC
---

# USE CASES

- Home theater automation
- TV / AV receiver control
- Audio system switching
- Macro automation
- Scripted remote control replacement
---

# SUMMARY

This system is a local IR automation platform consisting of:
- PowerShell frontend (command interface)
- Python Windows Service (execution engine)
- Named Pipe IPC communication
- USB IR hardware output layer
---

It enables software-based control of infrared devices through simple text commands.
