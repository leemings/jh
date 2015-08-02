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
rst.Open "select * from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
userhp=rst("体力")
usermp=rst("内力")
userat=rst("攻击")
userde=rst("防御")
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
<tr><td><input type=text value='生命:<%=userhp%>' size=12 disabled name=hp></td></tr>
<tr><td><input type=text value='内力:<%=usermp%>' size=12 disabled name=mp></td></tr>
<tr><td><input type=text value='攻击:<%=userat%>' size=12 disabled name=attack></td></tr>
<tr><td><input type=text value='防御:<%=userde%>' size=12 disabled name=defence></td></tr>
<tr><td><input type=button name=move value='移&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;动' onclick="javascript:parent.behfrm.location.href='action.asp'"></td></tr>
<tr><td><input type=button value='战&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;斗' onclick="javascript:parent.optfrm.location.href='attack.asp';parent.behfrm.location.replace('fight.asp')"></td></tr>
<tr><td><input type=button value='药&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;品' onclick="javascript:parent.optfrm.location.href='curative.asp'"></td></tr>
<tr><td><input type=button value='捕&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;捉' onclick="javascript:parent.actfrm.location.href='catch.asp';"></td></tr>
<tr><td><input type=button value='重&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;载' onclick='javascript:top.location.reload();'></td></tr>
<tr><td></td></tr>
</table>
</form>
</body>
<%

%>