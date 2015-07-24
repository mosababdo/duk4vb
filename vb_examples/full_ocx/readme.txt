
note: 

  --> this conversion to build as an ocx is still in early development ! <--


the debugger interface has several ActiveX control dependancies
which must be registered on your system for this to run.
(from a 32bit process) 

Register these before opening the project in VB6.

scivb_lite.ocx[1] - 410k, open source VB6 ActiveX control
SciLexer.dll  [2] - 460k, open source C standard dll
MSCOMCTL.OCX  [3] -   1Mb,closed source free from MS
Duk4VB.dll    [4] - 400k, (release mode) open source C standard dll
--------------------------------------
total dependancies: ~2mb

[1] https://github.com/dzzie/scivb_lite
[2] http://www.scintilla.org/  (build included w/scivb)
[3] probably pre-installed on system - from Microsoft 
[4] https://github.com/dzzie/duk4vb
    authors site: http://duktape.org/

Also note, duktape is capable of loading and debugging across 
multiple files. For my implementation I have chosen to only allow
the user to debug the active file they are working on. 

My personal use requirements will be to include other files
as libraries and com wrappers, but these arent something I want
to bother the user with having to see. If the user tries to single 
step into a library function, it will just stepout automatically.

You can change this around however you wish. 