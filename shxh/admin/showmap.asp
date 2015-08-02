<%
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
allright=Application("Ba_jxqy_allright")
if not(usergrade=allright and usercorp="官府") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from map order by posr,posc",conn
for i=1 to 8
	msg=msg&"<tr align=center><td bgcolor=ffff00>r"&i&"</td>"
	for j=0 to 11
	id=rst("id")
	posr=rst("posr")+1
	posc=rst("posc")+1
	displace=rst("displace")
	affair=rst("affair")
	msg=msg&"<td><a href='chgmap.asp?id="&id&"' title="&chr(34)&"R:"&posr&"；C:"&posc&"；可移动："&displace&"；事件："&affair&chr(34)&"><img src='../images/land/land_r"&posr&"_c"&posc&".gif' border=0></a></td>"
	rst.MoveNext
	next
	msg=msg&"</tr>"
next
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center>地图系统</p>
<hr>
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>r\c</td><td>c1</td><td>c2</td><td>c3</td><td>c4</td><td>c5</td><td>c6</td><td>c7</td><td>c8</td><td>c9</td><td>c10</td><td>c11</td><td>c12</td></tr>
<%=msg%>
</table>
</body>