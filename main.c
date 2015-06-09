#include "./duk/duktape.h"
#include <conio.h>
#include "vb.h"
#include <windows.h>

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

char* mLastString = 0;

enum opDuk{
	opd_PushUndef = 0,
	opd_PushNum =1,
	opd_PushStr =2,
	opd_GetInt=3,
	opd_IsNullUndef=4,
	opd_GetString=5,
	opd_Destroy=6,
	opd_LastString=7,
	opd_ScriptTimeout=8
};

int watchdogTimeout = 0;
unsigned int startTime = 0;
unsigned int lastRefresh = 0;

int __stdcall DukOp(int operation, duk_context *ctx, int arg1, char* arg2){
#pragma EXPORT
	
	//these do not require a context..
	switch(operation){
		case opd_LastString: return mLastString;
		case opd_ScriptTimeout: watchdogTimeout = arg1; return 0;
	}

	if(ctx == 0) return -1;

	switch(operation){
		case opd_PushUndef: duk_push_undefined(ctx); return 0;
		case opd_PushNum: duk_push_number(ctx,arg1); return 0;
		case opd_PushStr: duk_push_string(ctx, arg2); return 0;
		case opd_GetInt: return duk_to_number(ctx, arg1);
		case opd_IsNullUndef: return (int)duk_is_null_or_undefined(ctx, arg1);
		case opd_GetString: return (int)duk_safe_to_string(ctx, arg1);
		case opd_Destroy: duk_destroy_heap(ctx); ctx = 0; return 0;
	}

	return -1;

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


int ScriptTimeoutCheck(const void*udata)
{
	unsigned int tick = GetTickCount();

	if(vbStdOut){
		if(tick - lastRefresh > 500){
			vbStdOut(cb_Refresh,0);
			lastRefresh = tick;
		}
	}

    if (watchdogTimeout) {
        if (tick - startTime > watchdogTimeout)  return 1;
    }
    return 0;
}

static void sandbox_fatal(duk_context *ctx, duk_errcode_t code, const char *msg) {

	char buf[255]={0};
	char *def = "no message";

	if(msg==0) msg = def;

	if(strlen(msg) < 200){
		sprintf(buf, "%ld: %s", (long)code, msg);
		if(vbStdOut){
			vbStdOut(cb_Fatal, buf);
		}else{
			MessageBox(0,buf,"Fatal Error In DukTape",0);
		}
	}else{
		if(vbStdOut){
			vbStdOut(cb_Fatal, msg);
		}else{
			MessageBox(0,msg,"Fatal Error In DukTape",0);
		}
	}
	
	exit(1);  /* must not return */
}


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
	int realArgCount = 0;

	int n = duk_get_top(ctx);  //padded number of args..not usable for us here..
	
	if(n < 2) return 0; //we require at least 2 args for this function..
	if(vbHostResolver==NULL) return 0; 
	
	/*
	duk_push_this(ctx);
	duk_push_string(ctx, "this.arguments");
	duk_get_var(ctx);
	duk_push_string(ctx, "length");
	duk_get_prop(ctx,-1);
	meth = duk_safe_to_string(ctx, 0);

	//duk_push_string(ctx, "length");
	if(duk_pcall_prop(ctx, 0, -1)==DUK_EXEC_SUCCESS){
		
	}
	i = duk_to_number(ctx,0);
	i=0;
	*/

	//duk_eval_string(ctx, "this.arguments.length");
	//meth = duk_safe_to_string(ctx, -1);

	/*duk_push_string(ctx, "arguments");
	duk_push_string(ctx, "length");
    duk_call_prop(ctx, -1, 1);
	meth = duk_safe_to_string(ctx, -1);

	duk_get_prop_string(ctx, -2, "arguments");
	meth = duk_safe_to_string(ctx, -1);
	//duk_push_object(ctx);
	duk_get_prop_string(ctx, 0, "length");
	realArgCount = duk_to_number(ctx,-1);
	duk_pop(ctx);
	duk_pop(ctx);
	*/

	meth = duk_safe_to_string(ctx, 0);   //arg0 is obj.method string
	realArgCount = duk_to_number(ctx,1); //arg1 is arguments.length
	hasRetVal = vbHostResolver(meth, ctx, realArgCount);

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

void RegisterNativeHandlers(duk_context *ctx){
	
	duk_push_global_object(ctx);
	duk_push_c_function(ctx, comResolver, DUK_VARARGS);
	duk_put_prop_string(ctx, -2, "resolver");
	duk_pop(ctx);  /* pop global */

	duk_push_global_object(ctx);
	duk_push_c_function(ctx, prompt, 2);
	duk_put_prop_string(ctx, -2, "prompt");
	duk_pop(ctx);  /* pop global */


}

int __stdcall DukCreate(){
#pragma EXPORT
	//duk_context *ctx = duk_create_heap_default();
	duk_context *ctx = duk_create_heap(0, 0, 0, ScriptTimeoutCheck, sandbox_fatal);
	RegisterNativeHandlers(ctx);
	return (int)ctx;
}


int __stdcall AddFile(duk_context *ctx, char* pth){
#pragma EXPORT
	int rv;
	if(ctx == 0) return -1;
	startTime = GetTickCount();
	rv = duk_peval_file(ctx, pth); //0 = success
	if(rv != 0){
		setLastString(duk_safe_to_string(ctx, -1)); //error message..
	}
	duk_pop(ctx);  /* ignore result */
	return rv;
}

int __stdcall Eval(duk_context *ctx, char* js ){
#pragma EXPORT

	int rv = 0;

	if(ctx == 0) return -1;

	//safe to call eval to avoid fatal panic handler on syntax error
	duk_push_string(ctx, js);
	startTime = GetTickCount();
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
 



