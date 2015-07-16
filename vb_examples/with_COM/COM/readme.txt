
'this we dont support..
'return types string() or arrays in general
'methods which require object arguments
'functions with more than 10 args

These JavaScript files are used to create the structure in the JavaScript engine
for the COM objects we would like to use. Currently these are generated statically
by the user for each COM object they would like to add access to.

Once all of the corner cases are identified and stability is assured, then maybe
we will switch to on demand generation. For the foreseeable future however this is how
it is.

There are several limitations noted at the top currently. This should not be a problem
for the vast majority of methods though.

You do not have to generate a wrapper for every single method, only the ones you wish to use
this will keep the complexity lower and readability higher.

These wrappers are automatically generated using the COM bindings generator tool.
You simply paste in the VB 6 formatted function prototypes for the COM object
you wish to wrap in the top text box, specify the name of the object and hit parse.

I have been using the open source tlbviewer tool to extract all of the prototypes
for the COM objects. Compiled versions as well as source are available here:

https://github.com/dzzie/MAP

the require hInst checkbox, is used for COM objects that are returned from other methods.
Top level, objects do not require this. 

The format of the JavaScript files is to have a commented block of the VB 6 formatted prototypes
at the top. You can comment these out with leading ; or # characters if you do not want to enable
them, (even if they are included in the JavaScript)

When add object is called it will parse the prototypes and then load the JavaScript file
in the script engine. If the prototypes return any other object types, it will also require
those to be loaded at the same time. This will happen automatically. If they are not found you will get
an error message. You can disable problematic methods if required as mentioned above.

You will see prototypes such as the following

function fsoClass(){

	this.BuildPath = function(Path, Name){
		return resolver('fso.BuildPath', arguments.length,0, Path, Name);
	}

The first three arguments are always required by the resolver function. They map to the following C code: 

int comResolver(duk_context *ctx) {
	
	meth = duk_safe_to_string(ctx, 0);   //arg0 is obj.method string
	realArgCount = duk_to_number(ctx,1); //arg1 is arguments.length
	hInst = duk_to_number(ctx,2);       //arg2 is this com objects hinst variable if not a top level obj (0 if not)
	hasRetVal = vbHostResolver(meth, ctx, realArgCount, hInst);


You can also grant access to methods and properties on your own forms. Below is the prototypes for
my test form, giving JavaScript access to the text box from the js form.Text1.text = 

form.js:
-----------------------------------------
/* 
     property get Text1 as TextBox
*/

function FormClass(){
	//dummy
}

FormClass.prototype = {
	  
	  get Text1(){
			return resolver("form.Text1.get", 0, 0); 	
	  }
	  
};

var form = new FormClass();


When this JavaScript is loaded, and the prototype parsed it will then look for textbox.js 
and load that.


textbox.js:
-----------------------------------------
/* bindings for a vb textbox control

     property get Text as string
     property let Text as string 

*/

function textboxClass(){
	this.hInst=0;
}

textboxClass.prototype = {
	get Text (){
		return resolver("textbox.Text.get", 0, this.hInst); 	
	},
	set Text (val){
		resolver("textbox.Text.let", 1, this.hInst, val); 
	}
};





