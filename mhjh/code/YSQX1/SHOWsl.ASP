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
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/mj.asp")
rst.Open "商店",conn
do while not rst.EOF
id=rst("id")
baoname=rst("物品名")
gongji=rst("类型")
nazou=rst("说明")
huobao=rst("功能")
shijian=rst("精力")
fyu=rst("攻击")
tx=rst("防御")
dj=rst("特效")
hy=rst("会员")
tu=rst("银两")
msg=msg&"<tr><td>"&id&"</td><td>"&baoname&"</td><td>"&gongji&"</td><td>"&nazou&"</td><td>"&huobao&"</td><td>"&shijian&"</td><td>"&fyu&"</td><td>"&tx&"</td><td>"&dj&"</td><td>"&hy&"</td><td>"&tu&"</td><td><a href=chgsl.asp?id="&id&"&opt=修改>改</a>|<a href=chgsl.asp?id="&id&"&opt=删除>删</a></td></tr>"
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
<body oncontextmenu=self.event.returnValue=false topmargin=10 background='../chatroom/bg1.gif'>
<p align=center>飞星寻龙－龙粮管理</p>
<hr><br>
当武功类型显示为“轩”表明该项为轩辕武功，当为“佛”标明为出家人修炼的武功，当为“鬼”表明为鬼魂修炼的武功！
<table border=0 width=100% align=center>
<tr bgcolor=#FFFF00 align=center><td>ID</td><td>物品名</td><td>类型</td><td>说明</td><td>功能</td><td>精力</td><td>攻击</td><td>防御</td><td>特效</td><td>会员</td><td>银两</td><td><a href="chgsl.asp?opt=新增&amp;id=-1">新增</a></td></tr>
<%=msg%>
</table>
</body>