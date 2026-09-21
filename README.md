# GDID Tester utility

A small program in VB6/TwinBasic (nothing special) that allows you to view the GDID value and remove it if you want to stymie a point of tracking by MicroSoft. The program is very slimline, it is 32bit but there is a twinproj file in the source that you can use to compile to 64bits using TwinBasic. Note that the TwinBasic version will lag behind the VB6 version as the VB6 code is the 'master'.

<img width="531" height="381" alt="image" src="https://github.com/user-attachments/assets/88f14ed3-7bbd-4df7-9e97-db98bec7155d" />

**Fig. 01 The GDID Tester Utility**

HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid"

**This key is used to identify your PC and what it accesses on the net.**

The first time you run the program, the original GDID key will be extracted and should be visible by default. The program stores the original GDID elsewhere, the first time it runs. It then allows you to view and manipulate the actual GDID as known to Windows in order to change or blank it - to avoid tracking.

* Change the run interval by moving the **interval slider**. Setting the slider to 0 secs will stop the testing timer.

* If the **Enable Regular Testing** check box is ticked, the program will check the above key value every X seconds (according to slider) to see if the GDID has been repopulated. You will see that this occurs infrequently but regularly. For example, on system startup or resume from sleep, 
  if you use Edge or visit any MS site that accesses login.live.com, (Microsoft account, Store, OneDrive, Microsoft 365, account-linked UWP apps) then this value may well be re-populated with the same original GDID within a few minutes.
  Your local PC contains the cached version, the permanent version is stored on Microsoft's sites.

* The **Remove button** wipes the GDID manually.

* If the **Enable Automatic Removal** check box is ticked, then the program will wipe the GDID field immediately it is found to be populated. The Regular Testing checkbox will be ticked automatically.

* If the **Automatic Replacement** check box is enabled, the program will automatically regenerate a unique 64bit (16char) GDID whenever login.live or similar changes the local GDID key.
  This will obfuscate the GDID removing from potential trackers of your web usage, one less element to track.
  Note: The program will always create a new GDID when it first runs but once it has done that, the GDID will remain.  It will not generate a new GDID when the current GDID is blank.

* If the **Enable At Windows Startup** check box is ticked, then the program will start automatically when your windows o/s restarts.

All the above settings will be saved and restored on program restart.

* The red **'X'** button will clear the date/time drop down log showing when the GDID was changed.

* The **Read GDID** button will allow you to read the GDID that currently exists within the registry at the above key.

* The **Generate** button will allow you to test automatic removal by generating a unique 64bit (16char), a completely random GDID. The automatic replacement can also be tested.

* The **View log** will open the change log text file to show the dates and times that the GDID has been changed, either by remote modification or by local auto-generation using this tool.

* The **Dismiss** button closes the utility.

* The utility writes to a logfile to permanently store the GDID change date and times so that you can see what has happened to the GDID over time. The logfile is called GDIDChangeLog.log
  and it sits in the following folder:
  
  **C:\USERS\<username>\APPDATA\ROAMING\GDITester**

<img width="669" height="710" alt="image" src="https://github.com/user-attachments/assets/667474b1-1a6e-4997-b43d-a8fbaf696f0c" />

**Fig. 02 The logfile showing typical contents.** (Fake GDIDs)

This utility will allow you to see when the ID changes by some unknown use of an MS live service or by similar access by a tool you are inadvertently running within windows. 
It will change the GDID to something random to help prevent tracking on the web. This utility can make a change within 1 second of a change being made. 

Please be aware that changing the GDID you may block access to Windows 'apps' available from the Windows store and may prevent access to any Microsoft supplied services via the login.live service.
If you are a Windows desktop user and do not use the old mobile-type apps, nor do you use the Windows store and Onedrive that require a Microsoft login, then this limitation will not affect you. Personally, I use Windows 10 and 
traditional desktop applications, Dropbox for storage and a local login. I do not use a Microsoft login, nor do I intend to. I do not want to be tracked.

Once the GDID has been obfuscated, to prevent further modification of the GDID you could also run a tool such as ["Windows Firewall Notifier"](https://github.com/wokhan/WFN) that can be configured to prevent any unwanted changes to the GDID by preventing login.live.com access to your computer.

For more information on the GDID tracking key visit this link : https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-it/

Later changes to come:

* Complete variable changes to make it fully 64bit 
* Use the API to test the GDID key instead of a timer.
* Phase II - WIP - deal with the GDID contained within the following set of keys:
  Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\IdentityCRL\Immersive\production\Token\
  Loop through all the keys there eg.

   {12E984BD-5803-4D78-9EFB-BED7B9212C26}

  extract the DeviceId, match it with the known original GDID and see which match.
  Change the DeviceID to match the new generated GDID

* Phase III - deal with the GDID values within the following set of keys:
  Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IdentityCRL\NegativeCache
  Loop through all the keys there eg.

  0018C013A05744F3_S-1-5-21-732211230-4157827500-48361523-1001
  
  extract the first 16 chars of the key name, match it with the known original GDID and see which match.
  If found, change the keyname to match the new generated GDID
