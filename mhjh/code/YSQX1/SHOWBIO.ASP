<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "biological",conn
do while not rst.EOF
id=rst("id")
biological=rst("biological")
picture=rst("picture")
hp=rst("hp")
attack=rst("attack")
defence=rst("defence")
encourage=rst("encourage")
msg=msg&"<tr><td>"&id&"</td><td>"&biological&"</td><td><img src='"&picture&"' width=50 height=50></td><td>"&hp&"</td><td>"&attack&"</td><td>"&defence&"</td><td>"&encourage&"</td><td><a href=chgbio.asp?id="&id&"&opt=�޸�>�޸�</a> | <a href=chgbio.asp?id="&id&"&opt=ɾ��>ɾ��</a> </td></tr>"
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
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'>
<p align=center>�������</p>
<hr>
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>ID</td><td>����</td><td>ͼƬ</td><td>����</td><td>����</td><td>����</td><td>����</td><td><a href=chgbio.asp?opt=����&id=-1>������ͼ</a></td></tr>
<%=msg%>
</table>
</body>
