<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<2 then 
s=2-sj 
Response.write "<script language='javascript'>alert ('我是剑侠之神，防止刷钱\n 请您等上["&s&"秒钟]再操作！');setTimeout('history.back();',1000);</script>"
Response.End 
end if 
Application.Lock 
Application("yx8_mhjh_dg")=now() 
Application.UnLock 
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select 银两,作用 from 私塾 where id="&id,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=024"
umoney=rst("银两")
uresult=rst("作用")
rst.Close
rst.Open "select id from 用户 where 姓名='"&username&"' and 银两>="&umoney,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
set rst=nothing
conn.Execute "update 用户 set "&uresult&",银两=银两-"&umoney&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
Response.Write "<head><title>私塾</title><link rel=stylesheet href='../style.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  background='../chatroom/bg1.gif'  oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>承惠"&umoney&"两，欢迎再来！<br>三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p></body>"
%>
