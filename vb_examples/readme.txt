

full_ocx is pretty much done at this point I think. binary compatability is
not yet set though. You may have to recompile hostform.exe and jsthing.exe
to get them to work with whatever the latest version ID of the ocx is.

the ocx, with_debug, and jsThing use the open source scivb2 control:

https://github.com/dzzie/scivb2

You could compile this into dukdbg ocx directly if you want but it adds 
allot of complexity to an already complex code base. 

copies of scivb2.ocx and SciLexer.dll are included in the \dependancies
folder. the ocx must be registered with regsvr32 from a _32bit_ process
before you can use it.









