<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
pid=Request.QueryString("pid")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
allright=Application("Ba_jxqy_allright")
if not(usergrade=allright and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ballot where pid="&pid,conn
do while not rst.EOF
	id=rst("id")
	pid=rst("pid")
	parti=rst("parti")
	vote=rst("vote")
	msg=msg&"<tr><td>"&pid&"</td><td>"&parti&"</td><td>"&vote&"</td><td><a href=chgblt.asp?pid="&pid&"&id="&id&"&opt=�޸�>�޸�</a> | <a href=chgblt.asp?pid="&pid&"&id="&id&"&opt=ɾ��>ɾ��</a></td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center>ͶƱϵͳ</p>
<hr>
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>ͶƱ��</td><td>ѡ��</td><td>��Ʊ</td><%if usergrade=allright and usercorp="�ٸ�" then%><td><a href=chgblt.asp?opt=����&id=-1&pid=<%=pid%>>����ͶƱѡ��</a></td><%end if%></tr>
<%=msg%>
</table>
<p align=center><input type=button value='����' onclick=javascript:location.href='showballot.asp';></p>
</body>