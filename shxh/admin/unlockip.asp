<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
ip=Request.QueryString("IP")
chatroomsn=Request.QueryString("chatroomsn")
chatroomname=Application("Ba_jxqy_systemname")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
unlockipright=Application("Ba_jxqy_unlockipright")
lockiplist=Application("Ba_jxqy_lockip")
username=session("Ba_jxqy_username")
mycorp=session("Ba_jxqy_usercorp")
mygrade=session("Ba_jxqy_usergrade")
if username="" then Response.redirect "../error.asp?id=016"
if mygrade>lockipright and mycorp="官府" then
	Application.Lock
	Application("Ba_jxqy_lockip")=replace(lockiplist,";"&ip&";",";")
	Application.UnLock 
else
	Response.Redirect "../error.asp?id=046"
end if
msg="<head><link rel='stylesheet' href='../chatroom/style1.css'>.</head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"'><div heigh=100% align=center valign=middle>解锁工作完成<br><a href="&chr(34)&"javascript:location.replace('onlinelist.asp?chatroomsn="&chatroomsn&"');"&chr(34)&">返回</a></div></body>"
Response.write msg
%>