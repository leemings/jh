<%@ LANGUAGE=VBScript%>
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("sjjh_name")=""  then response.end
nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
myid=session("sjjh_id")
	
%>
<% Session("usepro")=true %>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE><%=Application("sjjh_chatroomname")%>-抗击倭寇</TITLE>
<SCRIPT LANGUAGE=JavaScript>
<!--

id=70;

password=70;

mebrood=100;

tobrood=100;

mode=1;

win=20;

//-->
</SCRIPT>

<SCRIPT LANGUAGE=JavaScript>
<!--
var InternetExplorer = navigator.appName.indexOf("Microsoft") != -1;
// Handle all the the FSCommand messages in a Flash movie


function bbk_DoFSCommand(command, args) {
  var xzObj = InternetExplorer ? bbk : document.bbk;

if(command=="loaddata")
{
xzObj.SetVariable("/:id",id);
xzObj.SetVariable("/:password",password);
xzObj.SetVariable("/:mebrood",mebrood);
xzObj.SetVariable("/:tobrood",tobrood);
xzObj.SetVariable("/:mode",mode);
xzObj.SetVariable("/:win",win);
xzObj.SetVariable("/:newdata",1);

}
}
// Hook for Internet Explorer 
if (navigator.appName && navigator.appName.indexOf("Microsoft") != -1 && 
	  navigator.userAgent.indexOf("Windows") != -1 && navigator.userAgent.indexOf("Windows 3.1") == -1) {
	
}else{
alert('对不起您使用的是NC的浏览器将有一些功能不能使用!建议更换使用IE5浏览')
}
//-->
</SCRIPT>
<SCRIPT LANGUAGE=VBScript>
	on error resume next 
	Sub bbk_FSCommand(ByVal command, ByVal args)
	call bbk_DoFSCommand(command, args)
	end sub
</SCRIPT>
</HEAD>
<BODY bgcolor="#ffffff" leftmargin="0" topmargin="0">
<!-- URL's used in the movie--> <A HREF=http://www.35city.com></A> <!-- text used in the movie--> 
<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
 codebase="http://active.macromedia.com/flash2/cabs/swflash.cab#version=4,0,0,0"
 ID=bbk WIDTH=550 HEIGHT=400>
  <PARAM NAME=movie VALUE="9.swf"> <PARAM NAME=quality VALUE=autolow> <PARAM NAME=devicefont VALUE=true> <PARAM NAME=bgcolor VALUE=#ffffff> 
  <EMBED src="9.swf" quality=autolow devicefont=true bgcolor=#ffffff  WIDTH=550 HEIGHT=400 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED></OBJECT>

</BODY>
</HTML>