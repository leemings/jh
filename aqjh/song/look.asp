<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=Application("aqjh_chatroomname")%>-->点歌祝福</title>
<link href="CSS.CSS" rel="stylesheet" type="text/css">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%
dim chinannid
chinannid=request("id")
set rs = Server.CreateObject("adodb.recordset")
sql="select *  from music where ID="+cstr(chinannid)+""
rs.open sql,conn,1,1
%>
<table align="center">
<tr><td><font color=red><%=rs("name")%></font>为<font color=red><%=rs("toname")%></font>送出一首[<a href="#" onClick="window.open('Play.asp?id=<%=rs("id")%>','zhufu','resizable=no,width=280,height=220')"><font color=blue><%=rs("songname")%></font></a>]</td></tr>
<tr><td>祝：<%=rs("zhufu")%></td></tr>
</table>
</body>
</html>
