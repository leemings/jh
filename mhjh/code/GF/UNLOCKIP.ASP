<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
ip=Request.QueryString("IP")
chatroomsn=Request.QueryString("chatroomsn")
chatroomname=Application("yx8_mhjh_systemname")
lockiplist=Application("yx8_mhjh_lockip")
username=session("yx8_mhjh_username")
mycorp=session("yx8_mhjh_usercorp")
mygrade=session("yx8_mhjh_usergrade")
if username="" then Response.redirect "../error.asp?id=016"
if mygrade>unlockipright and mycorp="官府" then
Application.Lock
Application("yx8_mhjh_lockip")=replace(lockiplist,";"&ip&";",";")
Application.UnLock
else
Response.Redirect "../error.asp?id=046"
end if
msg="<head><link rel='stylesheet' href='../chatroom/css.css'>.</head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><div heigh=100% align=center valign=middle>解锁工作完成<br><a href="&chr(34)&"javascript:location.replace('onlinelist.asp?chatroomsn="&chatroomsn&"');"&chr(34)&">返回</a></div></body>"
Response.write msg
%>
