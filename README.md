# GDID
GDID Tester utility

A small program in VB6/TwinBasic (nothing special) that allows you to view the GDID value and remove it if you want to stymie a point of tracking by MicroSoft

<img width="508" height="291" alt="image" src="https://github.com/user-attachments/assets/3784c0c7-e37e-47d7-946e-44136412a776" />

HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid"

**This key is used to identify your PC and what it accesses on the net.**

The program is very slimline, tt is 32bit but there is a twinproj file in the source that you can use to compile to 64bits using TwinBasic. 

The first time you run it the GDID key should be visible by default.

The remove button wipes it locally.

If the Enable Regular Testing check box is ticked, the program will check the above key value every ten seconds to see if it has been repopulated. You will see that this occurs regularly, for example on system startup or resume from sleep.
If you use Edge or visit any MS site that accesses login.live.com, (Microsoft account, Store, OneDrive, Microsoft 365, account-linked UWP apps) then this value may be re-populated with the same GDID.
Your local PC contains the cached version, the permanent version is stored on Microsoft's sites.

If the Enable Automatic Removal check box is ticked, then the program will wipe the GDID field immediately it is found to be populated. This requires the Regular Testing checkbox to be ticked too.

If the Enable At Windows Startup check box is ticked, then the program will start automatically when your windows o/s restarts.

All the above settings will be saved and restored on program restart.

The red 'X' button will clear the date/time drop down log showing when the GDID was changed.

This utility will allow you to see when the ID changes by some unknown use of an MS live service or by similar access by a tool you are inadvertently running within windows. 
However, to block this tracking you will need to run a tool such as Windows Firewall Notifier to block any unwanted login.live.com access.

For more information on the GDID tracking key visit this link : https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-it/

Later changes:
* write to a logfile to permamently store the GDID change date and times
* use an API to test the GDID key instead of a timer
* new code to generate a unique 64bit (16char) GDID
* automatically regenerate a unique 64bit (16char) GDID when login.live changes the local cache of GDID in the registry obfuscating the GDID.
