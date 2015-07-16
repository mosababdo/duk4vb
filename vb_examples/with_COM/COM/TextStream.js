/*
	Property Get Line  As Long
	Property Get Column  As Long
	Property Get AtEndOfStream  As Boolean
	Property Get AtEndOfLine  As Boolean
	Function Read ( ByVal Characters  As Long )  As String
	Function ReadLine  As String
	Function ReadAll  As String
	Sub Write ( ByVal Text  As String )
	Sub WriteLine ( ByVal Text  As String )
	Sub WriteBlankLines ( ByVal Lines  As Long )
	Sub Skip ( ByVal Characters  As Long )
	Sub SkipLine
	Sub Close
*/

function TextStreamClass(){

	this.hInst = 0

	this.Read = function(Characters){
		return resolver('TextStream.Read', arguments.length, this.hInst, Characters);
	}

	this.ReadLine = function(){
		return resolver('TextStream.ReadLine', arguments.length, this.hInst);
	}

	this.ReadAll = function(){
		return resolver('TextStream.ReadAll', arguments.length, this.hInst);
	}

	this.Write = function(Text){
		return resolver('TextStream.Write', arguments.length, this.hInst, Text);
	}

	this.WriteLine = function(Text){
		return resolver('TextStream.WriteLine', arguments.length, this.hInst, Text);
	}

	this.WriteBlankLines = function(Lines){
		return resolver('TextStream.WriteBlankLines', arguments.length, this.hInst, Lines);
	}

	this.Skip = function(Characters){
		return resolver('TextStream.Skip', arguments.length, this.hInst, Characters);
	}

	this.SkipLine = function(){
		return resolver('TextStream.SkipLine', arguments.length, this.hInst);
	}

	this.Close = function(){
		return resolver('TextStream.Close', arguments.length, this.hInst);
	}

}

TextStreamClass.prototype = {
	get Line(){
		return resolver('TextStream.Line.get', 0, this.hInst);
	},

	get Column(){
		return resolver('TextStream.Column.get', 0, this.hInst);
	},

	get AtEndOfStream(){
		return resolver('TextStream.AtEndOfStream.get', 0, this.hInst);
	},

	get AtEndOfLine(){
		return resolver('TextStream.AtEndOfLine.get', 0, this.hInst);
	}
}

var TextStream = new TextStreamClass()

