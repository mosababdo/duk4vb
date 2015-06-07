
function dlgClass(){
	
	this.ShowOpen = function(filt,initDir,title,hwnd){ 
		return resolver("call:cmndlg:OpenDialog:long:[string]:[string]:[long]:r_string", filt,initDir,title,hwnd); 
	}	
	
}

function fsoClass(){
	this.ReadFile = function(fname){
		return resolver("call:fso:ReadFile:string:r_string", fname); 	
	}	
}

var fso2 = {
	OpenTextFile : function(FileName,IOMode,Create,Format){
		return resolver("call:fso2:OpenTextFile:string:[long]:[bool]:[long]:r_objTextStreamClass", FileName,IOMode,Create,Format); 	
	}
	
}

function TextStreamClass(){
	this.hInst =0;
	this.ReadAll = function(){
		return resolver("call:objptr:ReadAll:r_string", this.hInst); 	
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
    resolver("let:form:caption:string", str); 
  }, 
  
  get caption() {
    return resolver("get:form:caption:string"); 
  },
  
  ReadFile : function(fname){
		return resolver("call:fso:ReadFile:string:r_string", fname); 	
  },  
  
  ShowOpen : function(filt,initDir,title,hwnd){ 
		return resolver("call:cmndlg:OpenDialog:long:[string]:[string]:[long]:r_string", filt,initDir,title,hwnd); 
  }	

}


/* this works
form.ReadFile = function(fname){
		return resolver("call:fso:ReadFile:string:r_string", fname); 	
}
*/



