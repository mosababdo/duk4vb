/*
		Sub AddItem(Item As variant, Index as integer)
		Property get Enabled As Boolean
		Property let Enabled As Boolean
		Property get ListCount As Integer
		Sub Clear()
*/

function listboxClass(){

	this.hInst = 0

	this.AddItem = function(Item, Index){
		return resolver('listbox.AddItem', arguments.length, this.hInst, Item, Index);
	}

	this.Clear = function(){
		return resolver('listbox.Clear', arguments.length, this.hInst);
	}

}

listboxClass.prototype = {
	get Enabled(){
		return resolver('listbox.Enabled.get', 0, this.hInst);
	},

	set Enabled(val){
		return resolver('listbox.Enabled.let', 1, this.hInst, val);
	},

	get ListCount(){
		return resolver('listbox.ListCount.get', 0, this.hInst);
	}
}

var listbox = new listboxClass()