<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
id=Request.QueryString("id")
bid=Request.QueryString("bid")
username=session("Ba_jxqy_username")
if not (isnumeric(bid) and isnumeric(id)) then Response.Redirect "../error.asp?id=220"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 版块 where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=224"
hostlist=";"&rst("版主")&";"
if not (instr(hostlist,";"&username&";")<>0 or instr(Ba_hostname,";"&username&";")<>0) then Response.Redirect "../error.asp?id=231"
rst.Close
rst.Open "select * from bbs where 论坛="&bid&" and id="&id,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=223"
if rst("精华")=true then
	ess=False
else
	ess=True
end if		
rst.Close
set rst=nothing
conn.execute "update bbs set 精华="&ess&" where id="&id&" and 论坛="&bid
conn.close
set conn=nothing
%>
<html><head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('location.replace("showboard.asp?bid=<%=bid%>");',3000);</script></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>信息成功发送，三秒钟后自动返回</font></td></tr><tr align=center><td><a href='javascript:location.replace("showboard.asp?bid=<%=bid%>");'>返回</a></td></tr></table></body></html>