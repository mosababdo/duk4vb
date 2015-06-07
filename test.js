
function dlgClass(){
	
	this.ShowOpen = function(filt,initDir,title,hwnd){ 
		return resolver("call:cmndlg:OpenDialog:long:[string]:[string]:[long]:r_string", arguments.length, filt,initDir,title,hwnd); 
	}	
	
}

function fsoClass(){
	this.ReadFile = function(fname){
		return resolver("call:fso:ReadFile:string:r_string", arguments.length, fname); 	
	}	
}

var fso2 = {
	OpenTextFile : function(FileName,IOMode,Create,Format){
		//if (arguments.length < 4) alert(arguments.length) 
		//this gives us the actual number of args passed into the script function
		//duk_get_top(ctx) padds missing arguments with null/undef I need the raw value for this..
		return resolver("call:fso2:OpenTextFile:string:[long]:[bool]:[long]:r_objTextStreamClass", arguments.length, FileName,IOMode,Create,Format); 	
	}
	
}

function TextStreamClass(){
	this.hInst=0;
	this.ReadAll = function(){
		return resolver("call:objptr:ReadAll:r_string", arguments.length, this.hInst); 	
	}
}


/*
Function OpenTextFile(FileName As String, 
						[IOMode As IOMode = ForReading], 
						[Create As Boolean = False], 
						[Format As Tristate = TristateFalse]
) As TextStream

Function ReadAll() As String Member of Scripting.TextStream

how to call a method on  a specific instance of an object instead of 
static top level global objects like we have been?

*/


var cmndlg = new dlgClass();
var fso = new fsoClass();

var form = {
  set caption (str) {
    resolver("let:form:caption:string", arguments.length, str); 
  }, 
  
  get caption() {
    return resolver("get:form:caption:string", arguments.length); 
  },
  
  ReadFile : function(fname){
		return resolver("call:fso:ReadFile:string:r_string", arguments.length, fname); 	
  },  
  
  ShowOpen : function(filt,initDir,title,hwnd){ 
		return resolver("call:cmndlg:OpenDialog:long:[string]:[string]:[long]:r_string", arguments.length, filt,initDir,title,hwnd); 
  }	

}


/* this works
form.ReadFile = function(fname){
		return resolver("call:fso:ReadFile:string:r_string", arguments.length, fname); 	
}
*/



