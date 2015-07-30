

note that full_ocx and jsThing are still in development.

you will probably have to compile any exes in these two directories
yourself to ensure they are using the latest ocx version
(binary compatiability not yet set)

the ocx and with_debug use the open source scivb2 control:

https://github.com/dzzie/scivb2

You could compile this into dukdbg ocx directly if you want but it adds 
allot of complexity to an already complex code base. 

copies of scivb2.ocx and SciLexer.dll are included in the \dependancies
folder. the ocx must be registered with regsvr32 from a _32bit_ process
before you can use it.









