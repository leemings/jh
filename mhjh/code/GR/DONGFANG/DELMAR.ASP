<!--#include file="data.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
sql="select * from 洞房 where ID="&id&" and 申请人='"&username&"'"
 Set Rst=connt.Execute(sql)
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=044"
rst.Close
set rst=nothing
connt.Execute "delete * from 洞房 where 申请人='"&username&"' and id="&ID
connt.Close
set connt=nothing
%>
<head><title></title><LINK rel=stylesheet href='../../style.css'><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'  topmargin=150>
<p align=center>3秒钟后自动返回<br> <font color=FF0000>好的我们马上替你撤消请求</font></p>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>
