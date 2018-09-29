
# ViGlance SE2

ViGlance is my own implementation of the Windows 7 Taskbar for Windows XP and Vista. It completely replaces your native windows taskbar. Stacks windows to their associated process in 1 neat group, previews window contents, gives you the ability to pin your favourite applications and it comes with the Windows 7 Start Orb. 

## Background 

The third Lee-Soft project to be released. Initially created for the Windows 7 Transformation Pack by Lucifiar. Developed in early 2009

## Libraries

- [Windows Unicode API TypeLib](https://github.com/badcodes/vb6/blob/master/%5BInclude%5D/TypeLib/winu.tlb) - Windows API, stores all the API declarations
- [dseaman@uol.com.br GDI+ Type Library 1.05](http://www.vbaccelerator.com/home/VB/Type_Libraries/GDIPlus_Type_Library/article.asp)
- [Karl E. Peterson's - HookMe](http://vb.mvps.org/samples/HookMe/) - A clean and elegant means of subclassing 
- [Extended GDIPlusWrapper](https://github.com/lee-soft/GDIPlusWrapper) - Extended GDIPlusWrapper used for OOP GDIPlus

## Getting Started

- Don't use the "download ZIP" feature of GitHub it will convert the windows new line characters to linux new line characters and VB6 can't understand them. Clone the repository with the GitHub desktop app 
- Ensure you have Visual Basic 6.0(Service Pack 6) installed
- Grab the WinU and GDIPlus TLB - extract the TLBs and add as a reference to the project
- Grab the HookMe zip - extract the files (IHookSink.cls, MHookMe.bas) over the placeholder files (IHookSink.cls, MHookMe.bas) and disregard any other files
- ~~Grab the Extended GDIPlusWrapper library - extract contents to "GDIPlusWrapper" and follow instructions for creating the library~~
- Grab the release of the [GDIPlusWrapper library](https://github.com/lee-soft/GDIPlusWrapper/releases) and re-add it as a reference to this project
- Compile and enjoy


## Acknolwedgements

I have been unable to contact the original author of the vbAccelerator GDIPlusWrapper (steve@vbaccelerator.com). Permission to include his library here is pending and until it is approved I will not be able to include it here.

## Libraries

- [Windows Unicode API TypeLib](https://github.com/badcodes/vb6/blob/master/%5BInclude%5D/TypeLib/winu.tlb) - Windows API, stores all the API declarations
- [dseaman@uol.com.br GDI+ Type Library 1.05](http://www.vbaccelerator.com/home/VB/Type_Libraries/GDIPlus_Type_Library/article.asp)
- [Karl E. Peterson's - HookMe](http://vb.mvps.org/samples/HookMe/) - A clean and elegant means of subclassing 
- [Extended GDIPlusWrapper](https://github.com/lee-soft/GDIPlusWrapper) - Extended GDIPlusWrapper used for OOP GDIPlus
