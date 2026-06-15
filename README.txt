========================================================
 SQL SERVER LITE MONITOR - README
========================================================

WHAT'S IN THIS PACKAGE
-----------------------
  SQLMonitor.ps1        - The actual monitor (PowerShell/WinForms)
  Run-SQLMonitor.bat    - Run directly WITHOUT compiling (easiest)
  Build-SQLMonitor.bat  - Compile into a proper .exe file
  README.txt            - This file

QUICKSTART (NO INSTALL NEEDED)
-------------------------------
  Option A - Run directly (fastest, no compile):
    Double-click:  Run-SQLMonitor.bat
    The tool opens immediately as a window.

  Option B - Compile to .exe (for distribution):
    Double-click:  Build-SQLMonitor.bat
    Needs internet once to grab PS2EXE from PowerShell Gallery.
    Creates SQLMonitor.exe you can copy anywhere.

REQUIREMENTS
------------
  - Windows 10, 11, Server 2016, 2019, 2022 or later
  - PowerShell 5.1 (built into Windows - no install)
  - SQL Server 2012 or later
  - VIEW SERVER STATE permission on SQL Server
  - For job history: msdb access needed

HOW TO CONNECT
--------------
  1. Enter server name (e.g.  MYSERVER  or  MYSERVER\INSTANCE)
  2. Choose Windows Auth (uses your current Windows login)
     OR SQL Server Auth and enter username/password
  3. Click Connect
  4. Dashboard populates - auto-refreshes every 30 seconds

DASHBOARD TABS
--------------
  ⚙  Configuration
      - Best practice checks (Max Memory, MAXDOP, etc.)
      - Color coded: GREEN=OK, ORANGE=Warning, BLUE=Info
      - Index fragmentation health

  ⚡  Performance
      - Live CPU usage (SQL vs System vs Idle)
      - Active running queries with wait types
      - Top wait statistics (filtered - noise removed)
      - Blocking sessions highlighted in red

  💾  Backups & Jobs
      - All user databases backup status
      - NEVER/WARNING highlighted for missing backups
      - SQL Agent job history (pass/fail/duration)

  📊  Queries & I/O
      - Top 20 slowest queries since last SQL restart
      - Per-file disk I/O latency (color coded)
      - Read/Write latency over 20ms = orange, 50ms = red

PERMISSIONS NEEDED ON SQL SERVER
---------------------------------
  GRANT VIEW SERVER STATE TO [your_login];
  -- For job history:
  USE msdb; EXEC sp_addrolemember 'SQLAgentReaderRole', 'your_login';

REFRESH SETTINGS
-----------------
  Use the dropdown next to "Refresh Now" to change from:
  15 seconds / 30 seconds / 60 seconds / Off (manual only)

NOTES
------
  - No data is sent anywhere - all queries run locally
  - No registry changes, no Windows services installed
  - Safe to run from USB drive or network share
  - Works with named instances: SERVER\INSTANCENAME
  - Works with port: SERVER,1433

========================================================
