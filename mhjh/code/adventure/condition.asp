<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select 体力,内力,攻击,防御 from 用户 where 姓名='"&username&"'",conn
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
<body background='../chatroom/bg1.gif' topmargin=10 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<form name=form1>
<table border=3 align=center width="108">
<tr><td align=center bgcolor=FFFF00 width="94"><%=username%>　</td></tr>
<tr><td width="94">
  <input type=text value='生命:<%=userhp%>' size=15 disabled name=hp></td></tr>
<tr><td width="94">
  <input type=text value='内力:<%=usermp%>' size=15 disabled name=mp></td></tr>
<tr><td width="94">
  <input type=text value='攻击:<%=userat%>' size=15 disabled name=attack></td></tr>
<tr><td width="94">
  <input type=text value='防御:<%=userde%>' size=15 disabled name=defence></td></tr>
<tr><td width="94" align="center"><input type=button name=move value='移&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;动' onclick="javascript:parent.behfrm.location.href='action.asp'"></td></tr>
<tr><td width="94" align="center"><input type=button value='战&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;斗' onclick="javascript:parent.optfrm.location.href='attack.asp';parent.behfrm.location.replace('fight.asp')"></td></tr>
<tr><td width="94" align="center"><input type=button value='药&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;品' onclick="javascript:parent.optfrm.location.href='curative.asp'"></td></tr>
<tr><td width="94" align="center"><input type=button value='捕&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;捉' onclick="javascript:parent.actfrm.location.href='catch.asp';"></td></tr>
<tr><td width="94" align="center"><input type=button value='逃&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;跑' onclick="javascript:parent.actfrm.location.href='taopao.asp';"></td></tr>
<tr><td width="94" align="center"><input type=button value='清&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;屏' onclick="javascript:parent.msgfrm.location.href='Refresh.asp';"></td></tr>
</table>
</form>
</body>
<%

%>