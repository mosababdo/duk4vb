/* bindings for a vb textbox control

#requires hInst

	property get Text as string
	property let Text as string 

*/

function fartboxClass(){
	this.hInst=0;
}

fartboxClass.prototype = {
	get Text (){
		return resolver("fartbox.Text.get", 0, this.hInst); 	
	},
	set Text (val){
		resolver("fartbox.Text.let", 1, this.hInst, val); 
	}
};

//this next line allows you to use a AddObject(txtMyTextBox, "textbox") directly..
var fartbox = new fartboxClass();