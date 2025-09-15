# ZKTeco Attendance & Parent SMS Notification System

## 1. Overview

This system provides a complete solution for automating school attendance tracking and parent communication using ZKTeco biometric devices. It polls attendance data in real-time, logs it to a central database and Excel files, and instantly sends SMS notifications to parents/guardians when a student enters or leaves the school.

The system is designed to be robust, reliable, and easy for school administrators to manage.

### Key Features

*   **Real-Time SMS Alerts:** Instantly informs parents of their child's movements.
*   **Accurate IN/OUT Logic:** Intelligently determines if a punch is a "Check-In" or "Check-Out" based on the function key used on the device.
*   **Centralized Logging:** All attendance records are saved to a central MySQL database for easy reporting and analysis.
*   **Individual Device Reports:** Each device automatically generates its own daily Excel log file.
*   **Robust & Stable:** Designed to run 24/7 as background services, with automatic error handling and recovery attempts.
*   **Multi-Device Support:** The architecture supports multiple devices running in parallel, each with its own isolated process.

## 2. System Architecture

The system is deployed using a "one-script-per-device" model for maximum stability and isolation.

*   **Main Project Folder:** A central folder (e.g., `C:\NjoroGirlsSystem\`) contains the management scripts and shared resources.
*   **Device Folders:** Inside the main folder, each biometric device has its own dedicated sub-folder (e.g., `DeviceA\`, `DeviceB\`).
*   **Isolated Scripts:** Each device folder contains a copy of the main `zkteco.py` script, configured specifically for that device.
*   **Launcher Scripts:** The main folder contains batch files (`.bat`) and a VBScript (`.vbs`) to easily start, stop, and silently launch all services.

### Data Flow

1.  A student punches a ZKTeco device.
2.  The dedicated Python script for that device polls for the new log every 3 seconds.
3.  The script reads the `log.punch` code to determine if it was a Check-In (`0`) or Check-Out (`1`).
4.  It retrieves the student's name and parent's phone number from the ZKTime 5.0 Access Database (`att2000.mdb`).
5.  The event is recorded in two places:
    *   The central `attendance_systems` MySQL database.
    *   The local `attendance_log_[DeviceName].xlsx` file inside the device's folder.
6.  An SMS message with the correct action is sent to the parent/guardian via the Sematime API.

## 3. Installation & Setup

### Prerequisites

1.  **Server:** A Windows machine that is always on.
2.  **XAMPP:** Installed with Apache and MySQL running. The `attendance_systems` database must be created.
3.  **Python:** Python 3.x installed. The path to `python.exe` should be in the system's PATH environment variable.
4.  **Python Libraries:** Install all required libraries by running `pip install -r requirements.txt`.
5.  **ZKTime 5.0:** The official software must be installed and communicating with the devices.
6.  **Network:** All devices and the server must be on the same network.

### Setup Steps

1.  **Database Setup:**
    *   Using phpMyAdmin, create a database named `attendance_systems`.
    *   Run the provided `database_setup.sql` script to create the `attendance_logs` and `sms_logs` tables.

2.  **Folder Structure:**
    *   Create a main project folder, for example: `C:\NjoroGirlsSystem\`.
    *   Inside, create a sub-folder for each biometric device (e.g., `DeviceA`, `DeviceB`, `NonTeaching`).

3.  **Configure Scripts:**
    *   Copy the main `zkteco.py` script into each device sub-folder.
    *   **Edit each copy:** Open the `zkteco.py` file in each folder and modify the `DEVICES` list at the top to contain only the IP address and name for that specific device.

4.  **Place Launcher Files:**
    *   Copy `start_all_services.bat`, `stop_all_services.bat`, and `launch_silently.vbs` into the main project folder (`C:\NjoroGirlsSystem\`).
    *   Edit the paths inside these files to ensure they point to the correct locations of your device folders.

## 4. How to Use

### Starting the System

To start all polling services, simply double-click the **`start_all_services.bat`** file. This will open a separate console window for each device, allowing you to monitor their live output.

**To start the system silently in the background:**
Double-click the **`launch_silently.vbs`** file. No windows will appear, but the processes will be running. You can verify this in the Windows Task Manager.

### Stopping the System

To stop all running services safely, double-click the **`stop_all_services.bat`** file. This will find and terminate all the correct processes.

### Setting Up Automatic Startup

To ensure the system starts automatically when the server boots up (highly recommended):
1.  Open the Windows **Task Scheduler**.
2.  Create a new task.
3.  Set the **Trigger** to **"At startup"**.
4.  Configure the task to **"Run whether user is logged on or not"**.
5.  Set the **Action** to "Start a program" and point it to your **`launch_silently.vbs`** script.

## 5. File Manifest

```
/NjoroGirlsSystem/
│
├── DeviceA/
│   └── zkteco.py           # Script configured for Device A
│
├── DeviceB/
│   └── zkteco.py           # Script configured for Device B
│
├── NonTeaching/
│   └── zkteco.py           # Script configured for Non-Teaching staff device
│
├── start_all_services.bat  # Starts all scripts in visible windows
├── stop_all_services.bat   # Stops all running scripts
├── launch_silently.vbs     # Starts all scripts silently in the background
├── README.md               # This documentation file
└── requirements.txt        # List of Python libraries to install
```

## 6. Technical Details & Troubleshooting

*   **Database:** The system connects to a MySQL database named `attendance_systems`. Ensure the credentials in `db_config` within the script are correct.
*   **Access DB:** The script reads from the ZKTime 5.0 database at `C:\Program Files (x86)\ZKTeco\ZKTime5.0\att2000.mdb`. Ensure this path is correct.
*   **Logs:** Each script generates its own detailed log files inside its `logs/` sub-folder. If a script is not working, these log files are the first place to check for errors.
*   **Timeout Errors:** If the logs show a `ZKNetworkError: timed out`, it means the script could not reach the device. Check the device's power, network cable, and that you can `ping` its IP address from the server.
*   **PID File Exists Error:** If a script fails to start and logs this error, it means it did not shut down cleanly. Run `stop_all_services.bat` or manually delete the `.pid` file from the script's folder before restarting.

---