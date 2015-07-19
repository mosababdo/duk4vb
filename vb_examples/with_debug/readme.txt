
note:

the debugger interface has several ActiveX control dependancies
which must be registered on your system for this to run.
(from a 32bit process)

scivb_lite.ocx[1] - 410k, open source VB6 ActiveX control
SciLexer.dll  [2] - 460k, open source C standard dll
spSubclass.dll[3] -  56k, closed source but free and easy to replace
MSCOMCTL.OCX  [4] -   1Mb,closed source free from MS
Duk4VB.dll    [5] - 400k, (release mode) open source C standard dll
--------------------------------------
total dependancies: ~2mb

[1] https://github.com/dzzie/scivb_lite
[2] http://www.scintilla.org/  (build included w/scivb)
[3] http://sandsprite.com/CodeStuff/subclassSetup.exe
[4] from Microsoft probably pre-installed on system
[5] https://github.com/dzzie/duk4vb
    authors site: http://duktape.org/

