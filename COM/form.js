/* bindings for my test form

	property get caption as string 
	property let caption as string
	property get Text1 as TextBox
	property get List2 as ListBox
*/

function FormClass(){
	//dummy
}

FormClass.prototype = {
	  set caption (str) {
	    	resolver("form.caption.let", 1, 0, str); 
	  }, 
	  
	  get caption() {
	    	return resolver("form.caption.get", 0, 0); 
	  },
	  
	  get Text1(){
			return resolver("form.Text1.get", 0, 0); 	
	  },
	  
	  get List2(){
			return resolver("form.List2.get", 0, 0); 	
	  }
	  
};

var form = new FormClass();