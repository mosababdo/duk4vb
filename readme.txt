
This project is a build of the C DukTape JavaScript engine
to be usable for visual basic 6.

Compiled size in release mode should about 400 K when statically
linked. The project is at a good state right now, a release 
snapshot at this point will be taken.

The reason for this project, is that the Microsoft script control
is quite old at this point and does not support many of the new
JavaScript features. Even simple things like the following are 
unsupported:

str = 'test'; alert(str[1])

The Microsoft script control also does not support debugging,
which is a large weakness.

General support for COM object methods has also been added 
to this build. Not every possibility is currently supported,
however the vast majority should be good. 

Since the COM support is an add-on to the engine, JavaScript
class wrappers will have to be built and loaded for the objects
you wish to utilize. There is an automatic bindings generator
to help you with this task. 

Examples are provided in the COM subfolder along with some 
sample scripts. See the README in this subfolder for more 
details.

If you use this please submit new (tested) bindings for 
common objects to help everyone along.


The following test cases all currently working:

'    js = "1+2"
'    js = "alert(1+2)"
'    js = "while(1){;}"                 'timeout test
'    js = "prompt('text')"
'    js = "a='testing';alert(a[0]);"

'------------- vbdevkit tests ---------------------
'    js = "fso2.ReadFile('c:\\lastGraph.txt')"
'    js = "alert(dlg.OpenDialog(4))"
'    js = "pth = dlg.OpenDialog(4,'title','c:\\',0); fso2.ReadFile(pth)"
'--------------------------------------------------

'    js = "form.Text1.Text = 'test'"
'    js = "form.Text1.Text + ' read back in from javascript!'"
'    js = "form.caption = 'test!';alert(form.caption)"
'    js = "for(i=0;i<10;i++)form.List2.AddItem('item:'+i);alert('clearing!');form.List2.Clear()"
'    js = "var ts = fso.OpenTextFile('c:\\lastGraph.txt',1,true,0);v = ts.ReadAll(); v"         'value of v is returned from eval..
'    js = "var ts = fso.OpenTextFile('c:\\lastGraph.txt',1); v = ts.ReadAll();alert(v)"         '(default args test)


benchmarks:
     + get the duktape javascript engine working with vb6
     + wire a general proxy between js and COM objects
     + create a COM bindings generator from vb6 prototypes
     + classify the engine and simplify methods to mimic ms script control
     + timeout mechanism and fatal app handler
     + auto loading of dependancies
     + handling of object return types (form.text1.text = etc)
     - integrate with debugger protocol of engine
     - use scintinilla (scivb) control to provide optional IDE/debugger UI
     - wrap into an ActiveX control for easy use into other projects


note some optional test cases uses one of my own activex dlls that 
you probably wont have. These tests self disable if the library is not
detected. You can find an installer for them here:

http://sandsprite.com/CodeStuff/vbdevkit.exe



