For conversion to 64 bit.

The issues are:

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
o Change the target to win64, save, compile, it should then work.
