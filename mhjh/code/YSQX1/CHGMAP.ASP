<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from map where id="&id ,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
displace=rst("displace")
affair=trim(rst("affair"))
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'>
<p align=center>地图系统</p><hr>
事件可用命令为：<br>
1.msg 信息 例：msg 听说你很菜，则在玩家走到此处时会显示信息“听说你你很菜"<br>
2.fight 怪物名 例：fight 罗杰，则玩家走到此处将遇到罗杰，请确信罗杰存在于怪物表中<br>
3 exit 没有参数，玩家走到此处则退出<br>
<form action='upmap.asp' method=post id=form1 name=form1>
<table width='80%' align=center border=3>
<tr><td>ID(只读)</td><td><input type=text name=id readonly value=<%=id%>></td></tr>
<tr><td>可移动(逻辑)</td><td><input type=text name=displace value="<%=displace%>" size=5 maxlength=5></td></tr>
<tr><td>事件(文本)</td><td><input type=text name=affair value="<%=affair%>" size=50 maxlength=25></td></tr>
<tr align=center><td colspan=2><input type=submit  value=' 更 改 '> <input type=button value='返回' onclick=javascript:location.href='showmap.asp';></td></tr>
</table>
</form>
