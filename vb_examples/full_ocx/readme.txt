
note: 

  --> this conversion to build as an ocx is still in development ! <--

So this OCX has 2 public components. one is a usercontrol which is a full IDE
with syntax highlighting and debugger controls.

The second is a public CDukTape class that allows you to use just the js
engine part of the code, without having to host a form and UI associated
with the debugger. 

This build does include the full COM integrations. see the notes
in ./../with_COM/COM/readme.txt for more details.

the debugger interface has several ActiveX control dependancies
which must be registered on your system for this to run.
(from a 32bit process) 

Register these before opening the project in VB6.

scivb2.ocx    [1] - 410k, open source VB6 ActiveX control
SciLexer.dll  [2] - 460k, open source C standard dll (no reg required)
MSCOMCTL.OCX  [3] -   1Mb,closed source free from MS
Duk4VB.dll    [4] - 400k, (release mode) open source C standard dll (no reg required)
--------------------------------------
total dependancies: ~2mb

[1] https://github.com/dzzie/scivb2
[2] http://www.scintilla.org/  (build included w/scivb2)
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

----------------------------------------


The built in duk> command line textbox lets you execute script calls
while debugging. Breakpoints will not be hit in sub functions however
as per how duktape was designed with regards to run time evals. 

scripts are also locked while running, there is no chance for edit and
continue. but at least you can reset script variable values, test
function outputs at runtime and print variable values at runtime.

this textbox also supports a couple built in commands.

.objs  - lists objects added to script envirnoment
.libs  - lists library files added 
.bl    - lists breakpoints
.cls   - raises a dbgout(cls) event for host to clear output window