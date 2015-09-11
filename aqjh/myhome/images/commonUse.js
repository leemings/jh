//===========common methods
function parseSpace(s) {//js could not allow space as a field
    var _s=s.split(" ");
	return _s.join("¡¡");
}
function formatTime(timeNum){
   var d=new Date(timeNum*1000);
   var s=d.getYear()+ "/"; s+=(d.getMonth()+1)+ "/"; s+=d.getDate()+"/";
   s+= d.getHours() +":";  s+= d.getMinutes() +":";   s+= d.getSeconds();
   return s;
}
function formatDay(timeNum){
     var d=new Date(timeNum*1000);
   var s=d.getYear()+ "/"; s+=(d.getMonth()+1)+ "/"; s+=d.getDate();
   return s;
}

function shortenStr(str,len){
	if(str==null||str.length==null||str.length<1)return "¡¡";
	if(str.length>len){
		return str.substr(0,len-1)+"..";
	}else{
		return str;
	}
}
