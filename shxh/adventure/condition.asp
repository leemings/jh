<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
userhp=rst("����")
usermp=rst("����")
userat=rst("����")
userde=rst("����")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel=stylesheet href='../style.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' topmargin=10 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<form name=form1>
<table border=3 align=center>
<tr><td align=center bgcolor=FFFF00><%=username%></td></tr>
<tr><td><input type=text value='����:<%=userhp%>' size=12 disabled name=hp></td></tr>
<tr><td><input type=text value='����:<%=usermp%>' size=12 disabled name=mp></td></tr>
<tr><td><input type=text value='����:<%=userat%>' size=12 disabled name=attack></td></tr>
<tr><td><input type=text value='����:<%=userde%>' size=12 disabled name=defence></td></tr>
<tr><td><input type=button name=move value='��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��' onclick="javascript:parent.behfrm.location.href='action.asp'"></td></tr>
<tr><td><input type=button value='ս&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��' onclick="javascript:parent.optfrm.location.href='attack.asp';parent.behfrm.location.replace('fight.asp')"></td></tr>
<tr><td><input type=button value='ҩ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ʒ' onclick="javascript:parent.optfrm.location.href='curative.asp'"></td></tr>
<tr><td><input type=button value='��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;׽' onclick="javascript:parent.actfrm.location.href='catch.asp';"></td></tr>
<tr><td><input type=button value='��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��' onclick='javascript:top.location.reload();'></td></tr>
<tr><td></td></tr>
</table>
</form>
</body>
<%

%>