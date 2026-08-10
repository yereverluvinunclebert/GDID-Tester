# GDID Tester utility

A small program in VB6/TwinBasic (nothing special) that allows you to view the GDID value and remove it if you want to stymie a point of tracking by MicroSoft. The program is very slimline, it is 32bit but there is a twinproj file in the source that you can use to compile to 64bits using TwinBasic. 

<img width="508" height="348" alt="image" src="https://github.com/user-attachments/assets/17c188ef-579d-4637-a17f-d5b983aa838b" />

HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid"

**This key is used to identify your PC and what it accesses on the net.**

The first time you run the program, the original GDID key should be visible by default. The program stores the original GDID the first time it runs and then allows you to view and manipulate the GDID in order to obfuscate it to avoid tracking.

* The remove button wipes the GDID manually.

* If the Enable Regular Testing check box is ticked, the program will check the above key value every ten seconds to see if it has been repopulated. You will see that this occurs regularly, for example on system startup or resume from sleep.
  If you use Edge or visit any MS site that accesses login.live.com, (Microsoft account, Store, OneDrive, Microsoft 365, account-linked UWP apps) then this value may be re-populated with the same GDID.
  Your local PC contains the cached version, the permanent version is stored on Microsoft's sites.

* If the Enable Automatic Removal check box is ticked, then the program will wipe the GDID field immediately it is found to be populated. The Regular Testing checkbox will be ticked automatically.

* If the Automatic Replacement check box is enabled, the program will automatically regenerate a unique 64bit (16char) GDID whenever login.live or similar changes the local GDID key.
  This will obfuscate the GDID. Note: When the GDID has automatically regenerated once, that GDID will remain.  

* If the Enable At Windows Startup check box is ticked, then the program will start automatically when your windows o/s restarts.

All the above settings will be saved and restored on program restart.

* The red 'X' button will clear the date/time drop down log showing when the GDID was changed.

* The Read GDID button will allow you to read the GDID that currently exists within the registry at the above key.

* The Generate button will allow you to test automatic removal by generating a unique 64bit (16char), a completely random GDID. The automatic replacement can also be tested.

* The Dismiss button closes the utility.

This utility will allow you to see when the ID changes by some unknown use of an MS live service or by similar access by a tool you are inadvertently running within windows. 
It will change the GDID to something random to help prevent tracking on the web. However, to fully block tracking by GDID you may need to run a tool such as Windows Firewall Notifier to prevent any unwanted changes to the GDID 
via login.live.com access.

For more information on the GDID tracking key visit this link : https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-it/

Later changes to come:

* write to a logfile to permanently store the GDID change date and times.
* use an API to test the GDID key instead of a timer
