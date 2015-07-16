/* 
# note: these vb prototypes are required to be here..they are parsed by the vb loader 
# they are automatically generated using the generator and the prototypes in this comment block

	Function BuildPath ( ByVal Path  As String ,  ByVal Name  As String )  As String
	Function GetDriveName ( ByVal Path  As String )  As String
	Function GetParentFolderName ( ByVal Path  As String )  As String
	Function GetFileName ( ByVal Path  As String )  As String
	Function GetBaseName ( ByVal Path  As String )  As String
	Function GetExtensionName ( ByVal Path  As String )  As String
	Function GetAbsolutePathName ( ByVal Path  As String )  As String
	Function GetTempName  As String
	Function DriveExists ( ByVal DriveSpec  As String )  As Boolean
	Function FileExists ( ByVal FileSpec  As String )  As Boolean
	Function FolderExists ( ByVal FolderSpec  As String )  As Boolean
	Sub DeleteFile ( ByVal FileSpec  As String ,  ByVal Force  As Boolean )
	Sub DeleteFolder ( ByVal FolderSpec  As String ,  ByVal Force  As Boolean )
	Sub MoveFile ( ByVal Source  As String ,  ByVal Destination  As String )
	Sub MoveFolder ( ByVal Source  As String ,  ByVal Destination  As String )
	Sub CopyFile ( ByVal Source  As String ,  ByVal Destination  As String ,  ByVal OverWriteFiles  As Boolean )
	Sub CopyFolder ( ByVal Source  As String ,  ByVal Destination  As String ,  ByVal OverWriteFiles  As Boolean )
	Function CreateTextFile ( ByVal FileName  As String ,  ByVal Overwrite  As Boolean ,  ByVal Unicode  As Boolean )  As TextStream
	Function OpenTextFile ( ByVal FileName  As String ,  ByVal IOMode As Long ,  ByVal Create  As Boolean ,  ByVal Format As Long )  As TextStream
	Function GetStandardStream ( ByVal StandardStreamType As Long ,  ByVal Unicode  As Boolean )  As TextStream
	Function GetFileVersion ( ByVal FileName  As String )  As String
	
	# note that I have not included the following functions yet because I did not feel like
	# generating the wrappers for the return types. You could probably fake create folder
	# and have it work just by deleting the return type or changing to a sub
	#
	#Property Get Drives  As IDriveCollection
	#Function GetDrive ( ByVal DriveSpec  As String )  As IDrive
	#Function GetFile ( ByVal FilePath  As String )  As IFile
	#Function GetFolder ( ByVal FolderPath  As String )  As IFolder
	#Function GetSpecialFolder ( ByVal SpecialFolder As Long )  As IFolder
	#Function CreateFolder ( ByVal Path  As String )  As IFolder


*/

function fsoClass(){



	this.BuildPath = function(Path, Name){
		return resolver('fso.BuildPath', arguments.length,0, Path, Name);
	}

	this.GetDriveName = function(Path){
		return resolver('fso.GetDriveName', arguments.length,0, Path);
	}

	this.GetParentFolderName = function(Path){
		return resolver('fso.GetParentFolderName', arguments.length,0, Path);
	}

	this.GetFileName = function(Path){
		return resolver('fso.GetFileName', arguments.length,0, Path);
	}

	this.GetBaseName = function(Path){
		return resolver('fso.GetBaseName', arguments.length,0, Path);
	}

	this.GetExtensionName = function(Path){
		return resolver('fso.GetExtensionName', arguments.length,0, Path);
	}

	this.GetAbsolutePathName = function(Path){
		return resolver('fso.GetAbsolutePathName', arguments.length,0, Path);
	}

	this.GetTempName = function(){
		return resolver('fso.GetTempName', arguments.length,0);
	}

	this.DriveExists = function(DriveSpec){
		return resolver('fso.DriveExists', arguments.length,0, DriveSpec);
	}

	this.FileExists = function(FileSpec){
		return resolver('fso.FileExists', arguments.length,0, FileSpec);
	}

	this.FolderExists = function(FolderSpec){
		return resolver('fso.FolderExists', arguments.length,0, FolderSpec);
	}

	this.DeleteFile = function(FileSpec, Force){
		return resolver('fso.DeleteFile', arguments.length,0, FileSpec, Force);
	}

	this.DeleteFolder = function(FolderSpec, Force){
		return resolver('fso.DeleteFolder', arguments.length,0, FolderSpec, Force);
	}

	this.MoveFile = function(Source, Destination){
		return resolver('fso.MoveFile', arguments.length,0, Source, Destination);
	}

	this.MoveFolder = function(Source, Destination){
		return resolver('fso.MoveFolder', arguments.length,0, Source, Destination);
	}

	this.CopyFile = function(Source, Destination, OverWriteFiles){
		return resolver('fso.CopyFile', arguments.length,0, Source, Destination, OverWriteFiles);
	}

	this.CopyFolder = function(Source, Destination, OverWriteFiles){
		return resolver('fso.CopyFolder', arguments.length,0, Source, Destination, OverWriteFiles);
	}

	this.CreateTextFile = function(FileName, Overwrite, Unicode){
		return resolver('fso.CreateTextFile', arguments.length,0, FileName, Overwrite, Unicode);
	}

	this.OpenTextFile = function(FileName, IOMode, Create, Format){
		return resolver('fso.OpenTextFile', arguments.length,0, FileName, IOMode, Create, Format);
	}

	this.GetStandardStream = function(StandardStreamType, Unicode){
		return resolver('fso.GetStandardStream', arguments.length,0, StandardStreamType, Unicode);
	}

	this.GetFileVersion = function(FileName){
		return resolver('fso.GetFileVersion', arguments.length,0, FileName);
	}

}

var fso = new fsoClass()

