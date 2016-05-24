/*
	Function GetListviewData(Index As Long)
	Function ResetAlertCount()
	Function DebugLog(msg,  data)
	Sub alert(x)
	Function eval(x)
	Function t(x)
	Function Save2Clipboard(x)
	Function GetClipboard()
	Sub writeFile(path, data)
	Function HexDump(str,  hexOnly = 0)
	Function ReadFile(path)
	Function unescape(x)
	Function pound_unescape(x)
	Function HexString2Bytes(str)
	Function EscapeHexString(hexstr)
*/

function tbClass(){



	this.GetListviewData = function(Index){
		return resolver('tb.GetListviewData', arguments.length,0, Index);
	}

	this.ResetAlertCount = function(){
		return resolver('tb.ResetAlertCount', arguments.length,0);
	}

	this.DebugLog = function(msg, data){
		return resolver('tb.DebugLog', arguments.length,0, msg, data);
	}

	this.alert = function(x){
		return resolver('tb.alert', arguments.length,0, x);
	}

	this.eval = function(x){
		return resolver('tb.eval', arguments.length,0, x);
	}

	this.t = function(x){
		return resolver('tb.t', arguments.length,0, x);
	}

	this.Save2Clipboard = function(x){
		return resolver('tb.Save2Clipboard', arguments.length,0, x);
	}

	this.GetClipboard = function(){
		return resolver('tb.GetClipboard', arguments.length,0);
	}

	this.writeFile = function(path, data){
		return resolver('tb.writeFile', arguments.length,0, path, data);
	}

	this.HexDump = function(str, hexOnly){
		return resolver('tb.HexDump', arguments.length,0, str, hexOnly = 0);
	}

	this.ReadFile = function(path){
		return resolver('tb.ReadFile', arguments.length,0, path);
	}

	this.unescape = function(x){
		return resolver('tb.unescape', arguments.length,0, x);
	}

	this.pound_unescape = function(x){
		return resolver('tb.pound_unescape', arguments.length,0, x);
	}

	this.HexString2Bytes = function(str){
		return resolver('tb.HexString2Bytes', arguments.length,0, str);
	}

	this.EscapeHexString = function(hexstr){
		return resolver('tb.EscapeHexString', arguments.length,0, hexstr);
	}

}

var tb = new tbClass()

