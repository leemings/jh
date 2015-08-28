<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("yx8_mhjh_userchatroomsn")
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"'",conn
for i=19 to 45
msg=msg&"<tr><td><font color=FF0000>"&rst.Fields(i).Name&"</font></td><td>"&rst.Fields(i).Value&"</td></tr>"
next
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<LINK href="css.css" rel=stylesheet>
</head>
<body oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5" background='bg1.gif'>
<div align=center>
基本状态<hr noshade size="1" color=red>
<table width=95% bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
<tr>
<td height="19" valign="top"><%=msg%></td>
</tr>
</table>
</div>
</body>
