/*  bindings for vbdevkit.clsFileSystem2

	Function FolderExists ( ByRef path  As String )  As Boolean
	Function FileExists ( ByRef path  As String )  As Boolean
	Function GetParentFolder ( ByRef path  As Variant )  As String
	Function CreateFolder ( ByRef path  As String )  As Boolean
	Function FileNameFromPath ( ByRef fullpath  As String )  As String
	Function WebFileNameFromPath ( ByRef fullpath  As String )
	Function DeleteFile ( ByRef fpath  As String )  As Boolean
	Sub Rename ( ByRef fullpath  As String ,  ByRef newName  As String )
	Sub SetAttribute ( ByRef fpath  As Variant ,  ByRef it As VBA.VbFileAttribute )
	Function GetExtension ( ByRef path  As Variant )  As String
	Function GetBaseName ( ByRef path  As String )  As String
	Function ChangeExt ( ByRef path  As String ,  ByRef ext  As String )
	Function SafeFileName ( ByRef proposed  As String )  As String
	Function RandomNum  As Long
	Function GetFreeFileName ( ByVal folder  As String ,  ByRef extension  As Variant )  As String
	Function GetFreeFolderName ( ByVal parentFolder  As String ,  ByRef prefix  As String )  As String
	Function buildPath ( ByRef folderPath  As String )  As Boolean
	Function ReadFile ( ByRef filename  As Variant )
	Sub WriteFile ( ByRef path  As String ,  ByRef it  As Variant )
	Sub AppendFile ( ByRef path  As Variant ,  ByRef it  As Variant )
	Function Copy ( ByRef fpath  As String ,  ByRef toFolder  As String )
	Function Move ( ByRef fpath  As String ,  ByRef toFolder  As String )
	Function CreateFile ( ByRef fpath  As String )  As Boolean
	Function DeleteFolder ( ByRef folderPath  As String ,  ByRef force  As Boolean )  As Boolean
	Function FolderName ( ByRef folderPath  As Variant )  As String
*/

function fso2Class(){
	
	this.FolderExists = function(path){
		return resolver('fso2.FolderExists', arguments.length, 0, path);
	}

	this.FileExists = function(path){
		return resolver('fso2.FileExists', arguments.length, 0, path);
	}

	this.GetParentFolder = function(path){
		return resolver('fso2.GetParentFolder', arguments.length, 0, path);
	}

	this.CreateFolder = function(path){
		return resolver('fso2.CreateFolder', arguments.length, 0, path);
	}

	this.FileNameFromPath = function(fullpath){
		return resolver('fso2.FileNameFromPath', arguments.length, 0, fullpath);
	}

	this.WebFileNameFromPath = function(fullpath){
		return resolver('fso2.WebFileNameFromPath', arguments.length, 0, fullpath);
	}

	this.DeleteFile = function(fpath){
		return resolver('fso2.DeleteFile', arguments.length, 0, fpath);
	}

	this.Rename = function(fullpath, newName){
		return resolver('fso2.Rename', arguments.length, 0, fullpath, newName);
	}

	this.SetAttribute = function(fpath, it){
		return resolver('fso2.SetAttribute', arguments.length, 0, fpath, it);
	}

	this.GetExtension = function(path){
		return resolver('fso2.GetExtension', arguments.length, 0, path);
	}

	this.GetBaseName = function(path){
		return resolver('fso2.GetBaseName', arguments.length, 0, path);
	}

	this.ChangeExt = function(path, ext){
		return resolver('fso2.ChangeExt', arguments.length, path, 0, ext);
	}

	this.SafeFileName = function(proposed){
		return resolver('fso2.SafeFileName', arguments.length, 0, proposed);
	}

	this.RandomNum = function(){
		return resolver('fso2.RandomNum', arguments.length, 0);
	}

	this.GetFreeFileName = function(folder, extension){
		return resolver('fso2.GetFreeFileName', arguments.length, 0, folder, extension);
	}

	this.GetFreeFolderName = function(parentFolder, prefix){
		return resolver('fso2.GetFreeFolderName', arguments.length, 0, parentFolder, prefix);
	}

	this.buildPath = function(folderPath){
		return resolver('fso2.buildPath', arguments.length, 0, folderPath);
	}

	this.ReadFile = function(filename){
		return resolver('fso2.ReadFile', arguments.length, 0, filename);
	}

	this.WriteFile = function(path, it){
		return resolver('fso2.WriteFile', arguments.length, 0, path, it);
	}

	this.AppendFile = function(path, it){
		return resolver('fso2.AppendFile', arguments.length, 0, path, it);
	}

	this.Copy = function(fpath, toFolder){
		return resolver('fso2.Copy', arguments.length, 0, fpath, toFolder);
	}

	this.Move = function(fpath, toFolder){
		return resolver('fso2.Move', arguments.length, 0, fpath, toFolder);
	}

	this.CreateFile = function(fpath){
		return resolver('fso2.CreateFile', arguments.length, 0, fpath);
	}

	this.DeleteFolder = function(folderPath, force){
		return resolver('fso2.DeleteFolder', arguments.length, 0, folderPath, force);
	}

	this.FolderName = function(folderPath){
		return resolver('fso2.FolderName', arguments.length, 0, folderPath);
	}

}

var fso2 = new fso2Class()

