<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select ����,����,����,���� from �û� where ����='"&username&"'",conn
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
<body background='../chatroom/bg1.gif' topmargin=10 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<form name=form1>
<table border=3 align=center width="108">
<tr><td align=center bgcolor=FFFF00 width="94"><%=username%>��</td></tr>
<tr><td width="94">
  <input type=text value='����:<%=userhp%>' size=15 disabled name=hp></td></tr>
<tr><td width="94">
  <input type=text value='����:<%=usermp%>' size=15 disabled name=mp></td></tr>
<tr><td width="94">
  <input type=text value='����:<%=userat%>' size=15 disabled name=attack></td></tr>
<tr><td width="94">
  <input type=text value='����:<%=userde%>' size=15 disabled name=defence></td></tr>
<tr><td width="94" align="center"><input type=button name=move value='��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��' onclick="javascript:parent.behfrm.location.href='action.asp'"></td></tr>
<tr><td width="94" align="center"><input type=button value='ս&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��' onclick="javascript:parent.optfrm.location.href='attack.asp';parent.behfrm.location.replace('fight.asp')"></td></tr>
<tr><td width="94" align="center"><input type=button value='ҩ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ʒ' onclick="javascript:parent.optfrm.location.href='curative.asp'"></td></tr>
<tr><td width="94" align="center"><input type=button value='��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;׽' onclick="javascript:parent.actfrm.location.href='catch.asp';"></td></tr>
<tr><td width="94" align="center"><input type=button value='��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��' onclick="javascript:parent.actfrm.location.href='taopao.asp';"></td></tr>
<tr><td width="94" align="center"><input type=button value='��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��' onclick="javascript:parent.msgfrm.location.href='Refresh.asp';"></td></tr>
</table>
</form>
</body>
<%

%>