<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
msg="<html><head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel=stylesheet href=../chatroom/css.css></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" ><table width=100% bgcolor=cccccc border=3><tr align=center><td colspan=2><font color=0000ff size=4>--"&showhomepage(Ba_homepage,Ba_pagename)&"--</font></td></tr>"
un=Request.QueryString("un")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&un&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=221"
msg=msg&"<tr bgcolor=f7f7f7><td width=40>��&nbsp;&nbsp;&nbsp;&nbsp;��</td><td>"&rst("����")&"</td></tr><tr><td>��&nbsp;&nbsp;&nbsp;&nbsp;��</td><td>"&rst("����")&"&nbsp;</td></tr><tr bgcolor=f7f7f7><td>��&nbsp;&nbsp;&nbsp;&nbsp;��</td><td>"&autolevel(rst("�ȼ�"))&"</td></tr><tr><td>��&nbsp;&nbsp;&nbsp;&nbsp;��</td><td>"&rst("�Ա�")&"</td></tr><tr bgcolor=f7f7f7><td>��������</td><td>"&showuseremail(rst("��������"))&"</td></tr><tr ><td>ǩ����</td><td>"
if rst("ǩ����")="" or isnull(rst("ǩ����")) then msg=msg&"δʹ��"
sign=rst("ǩ����")
msg=msg&Autolink(sign)&"</td></tr>"
rst.Close 
set rst=nothing
conn.close
set conn=nothing
msg=msg&"</table></body><html>"
Response.Write msg
%>