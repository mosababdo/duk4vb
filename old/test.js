
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


function TextBoxClass(){
	this.hInst=0;
}

TextBoxClass.prototype = {
	get Text (){
		return resolver("get:objptr:Text:r_string", arguments.length, this.hInst); 	
	},
	set Text (val){
		resolver("let:objptr:Text", arguments.length, this.hInst, val); 
	}
};

function FormClass(){
	//dummy
}

FormClass.prototype = {
	  set caption (str) {
	    	resolver("let:form:caption:string", arguments.length, str); 
	  }, 
	  
	  get caption() {
	    	return resolver("get:form:caption:string", arguments.length); 
	  },
	  
	  get Text1 (){
			return resolver("get:form:Text1:r_objTextBoxClass", arguments.length); 	
	  }
	  
};

var cmndlg = new dlgClass();
var fso = new fsoClass();
var form = new FormClass();

