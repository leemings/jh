<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("yx8_mhjh_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "x",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=5 background="../chatroom/bg1.gif">
<form method=post action=updatesystem.asp id=form1 name=form1>
<table width=100% border=3>
<tr><td width=30%>系统属性</td><td>属性值</td></tr>
<%
do while not rst.EOF%>
<tr><td><%=rst("a")%></td><td><input type=text name="<%=rst("a")%>" value="<%=rst("b")%>"  size=50 maxlength=300></td></tr>
<%
rst.MoveNext
loop%>
<tr><td colspan=2=2 align=center><input type=submit value=' 更 新' > <input type=reset value=' 重 置 '></td></tr>
</table>
</form>
</body>
<%
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>


