/* 
# bindings for vbdevkit.ClsCmnDlg2

	Property Let ErrorOnCancel  As Boolean
	Property Get ErrorOnCancel  As Boolean
	Sub SetCustomFilter ( ByRef displayText  As String ,  ByRef wildCardExtMatch  As Variant )
	Function OpenDialog ( ByRef filt As Long ,  ByRef initDir  As String ,  ByRef title  As String ,  ByRef pHwnd  As Long )  As String
	Function SaveDialog ( ByRef filt As Long ,  ByRef initDir  As String ,  ByRef title  As String ,  ByRef ConfirmOvewrite  As Boolean ,  ByRef pHwnd  As Long ,  ByRef defaultFileName  As String )  As String
	Function ColorDialog ( ByRef pHwnd  As Long )  As Long
	Function FolderDialog ( ByRef initDir  As String ,  ByRef pHwnd  As Long )  As String
*/

function dlgClass(){

	this.SetCustomFilter = function(displayText, wildCardExtMatch){
		return resolver('dlg.SetCustomFilter', arguments.length, displayText, wildCardExtMatch);
	}

	this.OpenDialog = function(filt, initDir, title, pHwnd){
		return resolver('dlg.OpenDialog', arguments.length, filt, initDir, title, pHwnd);
	}

	this.SaveDialog = function(filt, initDir, title, ConfirmOvewrite, pHwnd, defaultFileName){
		return resolver('dlg.SaveDialog', arguments.length, filt, initDir, title, ConfirmOvewrite, pHwnd, defaultFileName);
	}

	this.ColorDialog = function(pHwnd){
		return resolver('dlg.ColorDialog', arguments.length, pHwnd);
	}

	this.FolderDialog = function(initDir, pHwnd){
		return resolver('dlg.FolderDialog', arguments.length, initDir, pHwnd);
	}

}
/*
dlgClass.prototype = {
	set ErrorOnCancel(val){
		return resolver('dlg.ErrorOnCancel.let', arguments.length, val);
	},

	get ErrorOnCancel(){
		return resolver('dlg.ErrorOnCancel.get', arguments.length);
	}
}
*/
var dlg = new dlgClass()

