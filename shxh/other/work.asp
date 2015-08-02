<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from 打工 where id="&id,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=103"
ucondition=rst("条件")
uresult=rst("作用")
umoney=rst("银两")
rst.Close
rst.Open "select * from 用户 where 姓名='"&username&"' and "&ucondition,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=047"
rst.Close
set rst=nothing
conn.BeginTrans
conn.Execute "update 用户 set "&uresult&" where 姓名='"&username&"' and "&ucondition
conn.CommitTrans
conn.Close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel=stylesheet href='../chatroom/css.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>这里是"&umoney&"两银子，请收好了，欢迎再来！<br>三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p></body>"
%>