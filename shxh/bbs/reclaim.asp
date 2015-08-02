<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
if instr(Ba_hostname,";"&username&";")=0 then Response.Redirect "../error.asp?id=233"
bid=Request.QueryString("bid")
id=Request.QueryString("id")
if not (isnumeric(bid) and isnumeric(id))then Response.Redirect "../error.asp?id=220"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 父贴 from bbs where 删除=True and 论坛="&bid&" and id="&id,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=238"
fid=rst("父贴")
rst.Close
set rst=nothing
conn.execute "update bbs set 删除=False where id="&id&" and 论坛="&bid
if fid<>0 then
	conn.execute "update bbs set 回复=回复+1 where id="&fid&" and 论坛="&bid
end if	
conn.close
set conn=nothing
%>
<html><head><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>信息成功发送，三秒钟后自动关闭</font></td></tr><tr align=center><td><a href='#' onclick='javascript:window.close();'>关闭</a></td></tr></table></body></html>