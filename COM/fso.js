function fsoClass(){
	//Function BuildPath ( ByVal Path As String , ByVal Name As String ) As String
	this.BuildPath = function(Path, Name){
		return resolver('fso.BuildPath', arguments.length, Path, Name);
	}

	//Function GetDriveName ( ByVal Path As String ) As String
	this.GetDriveName = function(Path){
		return resolver('fso.GetDriveName', arguments.length, Path);
	}

	//Function GetParentFolderName ( ByVal Path As String ) As String
	this.GetParentFolderName = function(Path){
		return resolver('fso.GetParentFolderName', arguments.length, Path);
	}

	//Function GetFileName ( ByVal Path As String ) As String
	this.GetFileName = function(Path){
		return resolver('fso.GetFileName', arguments.length, Path);
	}

	//Function GetBaseName ( ByVal Path As String ) As String
	this.GetBaseName = function(Path){
		return resolver('fso.GetBaseName', arguments.length, Path);
	}

	//Function GetExtensionName ( ByVal Path As String ) As String
	this.GetExtensionName = function(Path){
		return resolver('fso.GetExtensionName', arguments.length, Path);
	}

	//Function GetAbsolutePathName ( ByVal Path As String ) As String
	this.GetAbsolutePathName = function(Path){
		return resolver('fso.GetAbsolutePathName', arguments.length, Path);
	}

	//Function GetTempName As String
	this.GetTempName = function(){
		return resolver('fso.GetTempName', arguments.length);
	}

	//Function DriveExists ( ByVal DriveSpec As String ) As Boolean
	this.DriveExists = function(DriveSpec){
		return resolver('fso.DriveExists', arguments.length, DriveSpec);
	}

	//Function FileExists ( ByVal FileSpec As String ) As Boolean
	this.FileExists = function(FileSpec){
		return resolver('fso.FileExists', arguments.length, FileSpec);
	}

	//Function FolderExists ( ByVal FolderSpec As String ) As Boolean
	this.FolderExists = function(FolderSpec){
		return resolver('fso.FolderExists', arguments.length, FolderSpec);
	}

	//Function GetDrive ( ByVal DriveSpec As String ) As IDrive
	this.GetDrive = function(DriveSpec){
		return resolver('fso.GetDrive', arguments.length, DriveSpec);
	}

	//Function GetFile ( ByVal FilePath As String ) As IFile
	this.GetFile = function(FilePath){
		return resolver('fso.GetFile', arguments.length, FilePath);
	}

	//Function GetFolder ( ByVal FolderPath As String ) As IFolder
	this.GetFolder = function(FolderPath){
		return resolver('fso.GetFolder', arguments.length, FolderPath);
	}

	//Function GetSpecialFolder ( ByVal SpecialFolder As Long ) As IFolder
	this.GetSpecialFolder = function(SpecialFolder){
		return resolver('fso.GetSpecialFolder', arguments.length, SpecialFolder);
	}

	//Sub DeleteFile ( ByVal FileSpec As String , ByVal Force As Boolean )
	this.DeleteFile = function(FileSpec, Force){
		return resolver('fso.DeleteFile', arguments.length, FileSpec, Force);
	}

	//Sub DeleteFolder ( ByVal FolderSpec As String , ByVal Force As Boolean )
	this.DeleteFolder = function(FolderSpec, Force){
		return resolver('fso.DeleteFolder', arguments.length, FolderSpec, Force);
	}

	//Sub MoveFile ( ByVal Source As String , ByVal Destination As String )
	this.MoveFile = function(Source, Destination){
		return resolver('fso.MoveFile', arguments.length, Source, Destination);
	}

	//Sub MoveFolder ( ByVal Source As String , ByVal Destination As String )
	this.MoveFolder = function(Source, Destination){
		return resolver('fso.MoveFolder', arguments.length, Source, Destination);
	}

	//Sub CopyFile ( ByVal Source As String , ByVal Destination As String , ByVal OverWriteFiles As Boolean )
	this.CopyFile = function(Source, Destination, OverWriteFiles){
		return resolver('fso.CopyFile', arguments.length, Source, Destination, OverWriteFiles);
	}

	//Sub CopyFolder ( ByVal Source As String , ByVal Destination As String , ByVal OverWriteFiles As Boolean )
	this.CopyFolder = function(Source, Destination, OverWriteFiles){
		return resolver('fso.CopyFolder', arguments.length, Source, Destination, OverWriteFiles);
	}

	//Function CreateFolder ( ByVal Path As String ) As IFolder
	this.CreateFolder = function(Path){
		return resolver('fso.CreateFolder', arguments.length, Path);
	}

	//Function CreateTextFile ( ByVal FileName As String , ByVal Overwrite As Boolean , ByVal Unicode As Boolean ) As ITextStream
	this.CreateTextFile = function(FileName, Overwrite, Unicode){
		return resolver('fso.CreateTextFile', arguments.length, FileName, Overwrite, Unicode);
	}

	//Function OpenTextFile ( ByVal FileName As String , ByVal IOMode As Long , ByVal Create As Boolean , ByVal Format As Long ) As ITextStream
	this.OpenTextFile = function(FileName, IOMode, Create, Format){
		return resolver('fso.OpenTextFile', arguments.length, FileName, IOMode, Create, Format);
	}

	//Function GetStandardStream ( ByVal StandardStreamType As Long , ByVal Unicode As Boolean ) As ITextStream
	this.GetStandardStream = function(StandardStreamType, Unicode){
		return resolver('fso.GetStandardStream', arguments.length, StandardStreamType, Unicode);
	}

	//Function GetFileVersion ( ByVal FileName As String ) As String
	this.GetFileVersion = function(FileName){
		return resolver('fso.GetFileVersion', arguments.length, FileName);
	}

}

fso.prototype = {
	//Property Get Drives As IDriveCollection
	get Drives(){
		return resolver('fso.Drives', arguments.length);
	}
}
