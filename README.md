﻿# Custom Windows Battery Notifcation Service

Custom made windows service. Provides a notification for low battery and fully charged (notify at 94%). Provides logging to Windows Event Log.
my inspo: nowhere, no one, just a pure frustration of wanting a notification for fully charged battery 🫠. Hence this quick custom notification service.

## Some Details

👉 .Net Framework 4.8

👉 C# lang version 10

## Steps to install the servie

1. Run vs 2022 or any compatible IDE as administrator.
2. Build the program either by pressing `Ctrl + Shift + b`, or right click on the solution in solution explorer, then click build.
3. Go to Developer Command Prompt, then go to CustomBatteryNotificationService/bin/Debug path. 
4. Type in `InstallUtil BatteryNotificationService.exe` to install the service, or `InstallUtil /u BatteryNotificationService.exe` to uninstall.
