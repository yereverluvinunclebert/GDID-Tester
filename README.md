# GDID Tester utility

A small program in VB6/TwinBasic (nothing special) that allows you to view the GDID value and remove it if you want to stymie a point of tracking by MicroSoft. The program is very slimline, it is 32bit but there is a twinproj file in the source that you can use to compile to 64bits using TwinBasic. 

<img width="531" height="381" alt="image" src="https://github.com/user-attachments/assets/88f14ed3-7bbd-4df7-9e97-db98bec7155d" />

HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid"

**This key is used to identify your PC and what it accesses on the net.**

The first time you run the program, the original GDID key should be visible by default. The program stores the original GDID the first time it runs and then allows you to view and manipulate the GDID in order to obfuscate it to avoid tracking.

* Change the run interval by moving the slider. Setting the slider to 0 secs will stop the testing timer.

* If the Enable Regular Testing check box is ticked, the program will check the above key value every ten seconds to see if it has been repopulated. You will see that this occurs regularly, for example on system startup or resume from sleep.
  If you use Edge or visit any MS site that accesses login.live.com, (Microsoft account, Store, OneDrive, Microsoft 365, account-linked UWP apps) then this value may be re-populated with the same GDID.
  Your local PC contains the cached version, the permanent version is stored on Microsoft's sites.

* The remove button wipes the GDID manually.

* If the Enable Automatic Removal check box is ticked, then the program will wipe the GDID field immediately it is found to be populated. The Regular Testing checkbox will be ticked automatically.

* If the Automatic Replacement check box is enabled, the program will automatically regenerate a unique 64bit (16char) GDID whenever login.live or similar changes the local GDID key.
  This will obfuscate the GDID. Note: The program will always create a new GDID when it first runs but once it has done that, the GDID will remain.  It will not generate a new GDID when the current GDID is blank.

* If the Enable At Windows Startup check box is ticked, then the program will start automatically when your windows o/s restarts.

All the above settings will be saved and restored on program restart.

* The red 'X' button will clear the date/time drop down log showing when the GDID was changed.

* The Read GDID button will allow you to read the GDID that currently exists within the registry at the above key.

* The Generate button will allow you to test automatic removal by generating a unique 64bit (16char), a completely random GDID. The automatic replacement can also be tested.

* The View log will open the change log text file to show the dates and times that the GDID has been changed, either by remote modification or by local auto-generation using this tool.

* The Dismiss button closes the utility.

* The utility writes to a logfile to permanently store the GDID change date and times so that you can see what has happened to the GDID over time. The logfile is called GDIDChangeLog.log
  and it sits in the following folder: C:\USERS\<username>\APPDATA\ROAMING\GDITester

This utility will allow you to see when the ID changes by some unknown use of an MS live service or by similar access by a tool you are inadvertently running within windows. 
It will change the GDID to something random to help prevent tracking on the web. This utility can make a change within 1 second of a change being made, however, to fully prevent and block tracking by GDID 
you could run a tool such as ["Windows Firewall Notifier"](https://github.com/wokhan/WFN) that can configured to prevent any unwanted changes to the GDID by preventing login.live.com access to your computer.

For more information on the GDID tracking key visit this link : https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-it/

Later changes to come:

* use an API to test the GDID key instead of a timer.
