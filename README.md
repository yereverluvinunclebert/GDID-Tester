# GDID
GDID Tester utility

A small program in VB6/TwinBasic (nothing special) that allows you to view the GDID value and remove it if you want to stymie a point of tracking by MicroSoft

<img width="485" height="257" alt="image" src="https://github.com/user-attachments/assets/90ae0b22-5c5a-4304-ad44-10cf96eb989c" />

HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid"

**This key is used to identify your PC and what it accesses on the net.**

The program is very slimline, the first time you run it the GDID key should be visible. It is 32bit but there is a twinproj file in the source that you can use to compile to 64bits using TwinBasic.

If enabled it checks the above key value every ten seconds to see if it has changed. You will see that it is repopulated regularly, for example on system startup or resume from sleep.
If you use Edge or visit any MS site that accesses login.live.com, (Microsoft account, Store, OneDrive, Microsoft 365, account-linked UWP apps) then this value may be re-populated with the same GDID.
Your local PC contains the cached version, the permanent version is stored on Microsoft's sites.

The remove button wipes it locally.

This utility will allow you to see when the ID changes by some unknown use of an MS live service or by similar access by a tool you are inadvertently running within windows. 

To block this tracking you will need to run a tool such as Windows Firewall Notifier to block any unwanted login.live.com access.

For more information on the GDID tracking key visit this link : https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-it/
