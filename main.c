#include "./duk/duktape.h"
#include <conio.h>


#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

#define real_fwrite fwrite

vbCallback vbStdOut;
vbHostResolverCallback vbHostResolver;

duk_context *ctx = 0 ; //vb6 is single threaded so lets simplify..
char* mLastString = 0;

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
	//vbLineInput = (vbDbgCallback)lpfnLineInput;
}

int setLastString(const char* s){
	if(mLastString != 0){ free(mLastString); mLastString = 0 ;}
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
	return setLastString(duk_safe_to_string(ctx, index));
}


int comResolver(duk_context *ctx) {
	int i, retType ;
	const char* meth = 0;

	int n = duk_get_top(ctx);  /* #args */

	if(n < 0) return 0;
	if(vbHostResolver==NULL) return 0; 
	
	meth = duk_safe_to_string(ctx, 0); //first arg is obj.method string
	retType = vbHostResolver(meth, strlen(meth), ctx, n-1);

	return 1;  
}

void RegisterNativeHandlers(){
	
	duk_push_global_object(ctx);
	duk_push_c_function(ctx, comResolver, DUK_VARARGS);
	duk_put_prop_string(ctx, -2, "resolver");
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
		vbStdOut( (fp == stdout ? cb_output : cb_error) , (char*)buf, sz);
	}
	else real_fwrite(buf, size, count, fp);
	
	return 0;
}
 
