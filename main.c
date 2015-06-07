#include "./duk/duktape.h"
#include <conio.h>
#include "vb.h"

/*
api tidbits:
	DUK_EXTERNAL_DECL duk_int_t duk_get_type(duk_context *ctx, duk_idx_t index);
	duk_is_none(), which would indicate whether index it outside of stack,
	is not needed; duk_is_valid_index() gives the same information.
*/


#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

#define real_fwrite fwrite

vbCallback vbStdOut = 0;
vbHostResolverCallback vbHostResolver = 0;
vbDbgCallback vbLineInput = 0;

duk_context *ctx = 0 ; //vb6 is single threaded so lets simplify..
char* mLastString = 0;

void __stdcall DukPushUndefined(duk_context *ctx){ 
#pragma EXPORT
	duk_push_undefined(ctx);
}

void __stdcall DukPushNum(duk_context *ctx, int num){ 
#pragma EXPORT
	duk_push_number(ctx,num);
}

void __stdcall DukPushString(duk_context *ctx, char* str){ 
#pragma EXPORT
	duk_push_string(ctx,str);
}

void __stdcall SetCallBacks(void* lpfnMsgHandler, void* lpfnDbgHandler, void* lpfnHostResolver, void* lpfnLineInput){
#pragma EXPORT
	vbStdOut     = (vbCallback)lpfnMsgHandler;
	//vbDbgHandler = (vbDbgCallback)lpfnDbgHandler;
	vbHostResolver = (vbHostResolverCallback)lpfnHostResolver;
	vbLineInput = (vbDbgCallback)lpfnLineInput;
}

int __stdcall setLastString(const char* s){ //accepts null string to just free last buffer
#pragma EXPORT
	if(mLastString != 0){ free(mLastString); mLastString = 0 ;}
	if(s==0) return 0;
	mLastString = strdup(s);
	return strlen(s);
}

int __stdcall GetLastStringSize(){ 
#pragma EXPORT
	if(mLastString == 0) return -1;
	return strlen(mLastString);
}

int __stdcall LastString(char* buf, int sz){
#pragma EXPORT
	if(mLastString){
		int a = strlen(mLastString);
		if(a < sz) strcpy(buf, (char*)mLastString);
		return a;
	}

	return -1;

}

int __stdcall DukGetInt(duk_context *ctx, int index){
#pragma EXPORT
	if(ctx == 0) return -1;
	return (int)duk_to_number(ctx, index);
}

//so even numeric args can be returned here like .ToString()
int __stdcall DukGetString(duk_context *ctx, int index){
#pragma EXPORT
	if(ctx == 0) return -1;
	//return setLastString(duk_safe_to_string(ctx, index));
	return (int)duk_safe_to_string(ctx, index);
}

/*
int __stdcall DukGethInst(duk_context *ctx, int index){
#pragma EXPORT
	int hInst=0;
	if(ctx == 0) return -1;
	duk_get_prop_string(ctx, index, "hInst"); 
	hInst = duk_to_number(ctx,index);
	return hInst;
}
*/

duk_ret_t js_dtor(duk_context *ctx)
{
	//int hInst = 0; hInst = duk_to_number(ctx, -1); 
	const char* hInst = 0;

    // The object to delete is passed as first argument
	//retrieve this.hInst numeric value
    duk_get_prop_string(ctx, 0, "hInst");
	hInst = duk_to_string(ctx, -1);
	if(vbStdOut!=0 && hInst!=0) vbStdOut(cb_ReleaseObj,hInst); 
	duk_pop(ctx); 
    return 0;

}

int __stdcall DukPushNewJSClass(duk_context *ctx, char* className, int hInst){
#pragma EXPORT
	if(ctx == 0) return -1;
	duk_get_global_string(ctx, className);
	duk_new(ctx, 0);

	if(hInst!=0){
		//its a live COM object instance we creating a js obj for..so set the hInst
		//and register a destructor function for it so we can cleanup when js no longer needs it..
		//set this.hInst = 12345
		duk_push_number(ctx, hInst); 
		duk_put_prop_string(ctx, -2, "hInst"); 

		//register a C function to run when js obj released
		duk_push_c_function(ctx, js_dtor, 1);
		duk_set_finalizer(ctx, -2);	 
	}

	return 0;

}


int comResolver(duk_context *ctx) {
	int i, hasRetVal ;
	const char* meth = 0;

	int n = duk_get_top(ctx);  /* #args */

	if(n < 0) return 0;
	if(vbHostResolver==NULL) return 0; 
	
	meth = duk_safe_to_string(ctx, 0); //first arg is obj.method string
	hasRetVal = vbHostResolver(meth, ctx, n-1);

	if(hasRetVal != 0 && hasRetVal != 1){
			MessageBox(0,"comresolver","vbdev the hasRetVal must be 0 or 1",0);
			hasRetVal = 1;
	}

	return hasRetVal;  
}

int prompt(duk_context *ctx){
	//prompt(text,defaultText)
	const char* meth = 0;
	int i, hasRetVal ;
	int n = duk_get_top(ctx);  /* #args */

	if(n < 0) return 0;
	if(vbLineInput==NULL) return 0; 
	
	meth = duk_safe_to_string(ctx, 0); 
	hasRetVal = vbLineInput(meth, ctx);
	return 1;
}

void RegisterNativeHandlers(){
	
	duk_push_global_object(ctx);
	duk_push_c_function(ctx, comResolver, DUK_VARARGS);
	duk_put_prop_string(ctx, -2, "resolver");
	duk_pop(ctx);  /* pop global */

	duk_push_global_object(ctx);
	duk_push_c_function(ctx, prompt, 2);
	duk_put_prop_string(ctx, -2, "prompt");
	duk_pop(ctx);  /* pop global */


}

void __stdcall DukDestroy(){
#pragma EXPORT
	duk_destroy_heap(ctx);
	ctx = 0;
}

void __stdcall DukCreate(){
#pragma EXPORT
	if(ctx != 0) DukDestroy();
	ctx = duk_create_heap_default();
	RegisterNativeHandlers();
}



int __stdcall AddFile(char* pth){
#pragma EXPORT
	int rv;
	if(ctx == 0) return -1;
	rv = duk_peval_file(ctx, pth); //0 = success
	if(rv != 0){
		setLastString(duk_safe_to_string(ctx, -1)); //error message..
	}
	duk_pop(ctx);  /* ignore result */
	return rv;
}



int __stdcall Eval(char* js){
#pragma EXPORT

	int rv = 0;

	if(ctx == 0) return -1;

	//safe to call eval to avoid fatal panic handler on syntax error
	duk_push_string(ctx, js);
	if (duk_peval(ctx) != 0) {
		setLastString(duk_safe_to_string(ctx, -1));
		rv = -1;
	} else {
		setLastString(duk_safe_to_string(ctx, -1));
	}

	duk_pop(ctx);
	return rv;

}


 
int my_fwrite( const void *buf, size_t size, size_t count, FILE* fp){
	
	int sz = size * count;

	if(vbStdOut != NULL && (fp == stdout || fp == stderr) ){
		vbStdOut( (fp == stdout ? cb_output : cb_error) , (char*)buf);
	}
	else real_fwrite(buf, size, count, fp);
	
	return 0;
}
 
