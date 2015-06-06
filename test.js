
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



