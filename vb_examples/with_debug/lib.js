
//this is to test behavior when running code in files added as libraries
//the behavior i want is to just step out of lib files..99% of the time in my 
//usage I wont want to debug these files and dont want the user to be bothered
//by it (such as the COM wrappers). also this makes the UI less complex since
//i dont have to show loaded files, buffer view switches etc.
function add_two(b){
	return b+=2
}