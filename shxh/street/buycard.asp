<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
cid=Request.QueryString("cid")
if not isnumeric(cid) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from cardshop where id="&cid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=029"
cname=rst("cname")
cespecial=rst("cespecial")
ctime=rst("ctime")
cprice=rst("cprice")
rst.Close
if Application("Ba_jxqy_fellow")=true then
	rst.Open "select 银两 from 用户 where 姓名='"&username&"' and 会员=true and 银两>="&cprice,conn
else
	rst.Open "select 银两 from 用户 where 姓名='"&username&"' and 银两>="&cprice,conn
end if	
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
rst.Open "select * from card where cname='"&cname&"' and username='"&username&"'",conn
conn.BeginTrans
if rst.EOF or rst.BOF then
	conn.Execute "insert into card(username,cname,cespecial,ctime,cnumber) values('"&username&"','"&cname&"','"&cespecial&"',"&ctime&",1)"
else
	conn.Execute "update card set cnumber=cnumber+1 where id="&rst("id")
end if
conn.Execute "update 用户 set 银两=银两-"&cprice&" where 姓名='"&username&"'"
conn.CommitTrans
rst.Close
set rst=nothing
conn.Close
set conn=nothing	
%>
<head><title>道具店</title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topMargin=200>
<div align=center>3秒钟后自动返回<br>
承惠纹银<%=cprice%>两，这是<%=cname%>，您收好了！
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>