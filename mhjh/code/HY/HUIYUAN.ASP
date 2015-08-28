<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
if username=""  then Response.Redirect "../error.asp?id=016"
if yx8_mhjh_fellow<>true then Response.Redirect "../error.asp?id=248"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 会员 from 用户 where 姓名='" & username & "'"
set rs=conn.execute(sql)
if rs("会员")=false then
%>
<script language=vbscript>
MsgBox "你不是会员，不能进会员战场打倭寇!"
window.close()
</script>
<%
conn.close
response.end
else
%>
<% Session("usepro")=true %>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>会员战场</TITLE>
<SCRIPT LANGUAGE=JavaScript>
<!--
id=70;
password=70;
mebrood=250;
tobrood=250;
mode=3;
win=25;

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
<!-- URL's used in the movie--> <A HREF=http://www.21ex.com></A> <!-- text used in the movie-->
<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
codebase="http://active.macromedia.com/flash2/cabs/swflash.cab#version=4,0,0,0"
ID=bbk WIDTH=550 HEIGHT=400>
<PARAM NAME=movie VALUE="../image/9.swf"> <PARAM NAME=quality VALUE=autolow> <PARAM NAME=devicefont VALUE=true> <PARAM NAME=bgcolor VALUE=#ffffff>
<EMBED src="../image/9.swf" quality=autolow devicefont=true bgcolor=#ffffff  WIDTH=550 HEIGHT=400 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED></OBJECT>

</BODY>
</HTML>
<%
conn.close
set rs=nothing
end if%>
