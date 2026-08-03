# GDID
GDID Tester utility

A small program in VB6/TwinBasic (nothing special) that allows you to view the GDID value and remove it if you want to stymie a point of tracking by MicroSoft

<img width="485" height="225" alt="image" src="https://github.com/user-attachments/assets/4be5fb3e-8a8c-48d6-b14c-eef7dcb41660" />

HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid"

**This key is used to identify your PC and what it access on the net.**

The program is very slimline, the first time you run it the GDID key should be visible

It checks the above value every ten seconds to see if it has changed.
If you use Edge or visit any MS site that accesses login.live.com, (Microsoft account, Store, OneDrive, Microsoft 365, account-linked UWP apps) then this value may be re-populated with a GDID.

The remove button wipes it.

This utility will allow you to see when the ID changes by some unknown use of MS live services or by access of a same by a tool you are inadvertently running within windows. 

To block this tracking you will need to run a tool such as Windows Firewall Notifier to block any unwanted login.live.com access.

For more information on the GDID tracking key visit this link : https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-it/
