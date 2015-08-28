<%
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then session.Abandon
umoney=rst("银两")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>悦来客栈</title><link rel="stylesheet" href="../style.css"></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false topmargin=50>
<p align=center><b><font color="#000000" face="方正舒体" size="4">悦来客栈</font></b></p>
<form action=sleep.asp method=post><table bgcolor=#ffffff border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFffff width=50% align=center><tr><td>收费标准：每六十两银子一个小时,你有现银<%=umoney%>两，可以体息<font color=ff0000><%=umoney\60%></font>个小时</td></tr><tr><td>我要体息<input type=text name=utime size=3 maxlength=3>个小时</td></tr><tr align=center><td><input type=submit value=' 睡 了 '> <input type=button onclick='javascript:window.close();' value=' 关 闭 '></td></tr></table></form>
<p align=center>
&nbsp;
</p>
<div align="center">
<center>
<table border="0" width="81%" cellspacing="0" cellpadding="0">
<tr>
<td width="20%" align="center"><img border="0" src="../image/h02.gif"></td>
<td width="20%" align="center"><img border="0" src="../image/h03.gif"></td>
<td width="20%" align="center"><img border="0" src="../image/h04.gif"></td>
<td width="20%" align="center"><img border="0" src="../image/h06.gif"></td>
<td width="20%" align="center"><img border="0" src="../image/h05.gif"></td>
<td width="20%" align="center"><img border="0" src="../image/h01.gif"></td>
</tr>
</table>
</center>
</div>
</body>
