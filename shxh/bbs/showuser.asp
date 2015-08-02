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
rst.Open "select * from 用户 where 姓名='"&un&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=221"
msg=msg&"<tr bgcolor=f7f7f7><td width=40>姓&nbsp;&nbsp;&nbsp;&nbsp;名</td><td>"&rst("姓名")&"</td></tr><tr><td>门&nbsp;&nbsp;&nbsp;&nbsp;派</td><td>"&rst("门派")&"&nbsp;</td></tr><tr bgcolor=f7f7f7><td>等&nbsp;&nbsp;&nbsp;&nbsp;级</td><td>"&autolevel(rst("等级"))&"</td></tr><tr><td>性&nbsp;&nbsp;&nbsp;&nbsp;别</td><td>"&rst("性别")&"</td></tr><tr bgcolor=f7f7f7><td>电子邮箱</td><td>"&showuseremail(rst("电子邮箱"))&"</td></tr><tr ><td>签名档</td><td>"
if rst("签名档")="" or isnull(rst("签名档")) then msg=msg&"未使用"
sign=rst("签名档")
msg=msg&Autolink(sign)&"</td></tr>"
rst.Close 
set rst=nothing
conn.close
set conn=nothing
msg=msg&"</table></body><html>"
Response.Write msg
%>