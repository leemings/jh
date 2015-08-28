<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=0
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
set rst=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/jihe.asp")
rst.Open "宝物",conn
do while not rst.EOF
id=rst("id")
baoname=rst("宝物名")
gongji=rst("攻击力")
nazou=rst("拿走")
huobao=rst("获宝者")
shijian=rst("时间")
fyu=rst("防御力")
jg=rst("价格")
msg=msg&"<tr><td>"&id&"</td><td>"&baoname&"</td><td>"&gongji&"</td><td>"&nazou&"</td><td>"&huobao&"</td><td>"&shijian&"</td><td>"&fyu&"</td><td>"&jg&"</td><td><a href=chgbao.asp?id="&id&"&opt=修改>修改</a> | <a href=CHGbao.ASP?id="&id&"&opt=删除>删除</a> </td></tr>"
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
<p align=center>海岛宝物管理</p>
<hr><br>
当拿走为1时，表明宝物被拿了，所以要把拿走改为0，这样玩家可以继续找到宝物了
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>ID</td><td>宝物名</td><td>攻击力</td><td>拿走</td><td>获宝者</td><td>时间</td><td>防御力</td><td>价格</td><td><a href=chgbao.asp?opt=新增&id=-1>新增</a></td></tr>
<%=msg%>
</table>
</body>