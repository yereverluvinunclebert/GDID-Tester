o  conversion to 64 bit.

The conversion issues are:

o The Krool slider.OCX is 32bit and therefore cannot work in conjunction with 64bit TB nor does it work with the potential TB fusion integration.
o The manifest required for the VB6 program to access the OCX refers to x86 architecture, incompatible with 64bit.

Fixes: 

o Transfer the code to the TB folder
o Import VBP
o Save to twinProj
o Open the mainfest
o Set the processor architecture from x86 to *
o Save
o Remove the ocx slider from the form using TB's form designer
o In Project References select packages, add Krools VBCCR18 package, common controls replacement
o Add a new slider in place of the other removed, with the same name.
o search for all occurrences of "required as longPtr", add the conditional compiler statements to use longPtr rather than a Long
o Add WinDevLib package to the program.
o Change the target to win64, save, compile, it should then work.

The GDIDTester cannot be converted fully to TB unless I can import a file into TB or export as I require. Also, cannot use source control whilst within a twinProj.
So, waiting until 1.0.

Tasks:

o Use the API to test the GDID key instead of a timer.

Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\IdentityCRL\Immersive\production\Token\
Loop through all the keys there eg.

{12E984BD-5803-4D78-9EFB-BED7B9212C26}

extract the DeviceId, match it with the known original GDID and see which match.
Change the DeviceID to match the new generated GDID

Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IdentityCRL\NegativeCache
Loop through all the keys there eg.
extract the first 16 chars of the key name, match it with the known original GDID and see which match.
0018C013A05744F3_S-1-5-21-732211230-4157827500-48361523-1001
If found, change the DeviceID to match the new generated GDID
stop the service

To stop the Connected Devices Platform User Service
HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\
Look for the service keys named CDPUserSvc and any corresponding split keys with alphanumeric suffixes such as CDPUserSvc_xxxx
4 = Disabled
