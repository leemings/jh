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
rst.Open "mj",conn
do while not rst.EOF
id=rst("id")
baoname=rst("名称")
gongji=rst("所有者")
nazou=rst("类型")
huobao=rst("攻击")
shijian=rst("防御")
fyu=rst("速度")
tx=rst("特效")
dj=rst("等级")
tu=rst("图")
sj=rst("时间")
zt=rst("状态")
msg=msg&"<tr><td>"&id&"</td><td>"&baoname&"</td><td>"&gongji&"</td><td>"&nazou&"</td><td>"&huobao&"</td><td>"&shijian&"</td><td>"&fyu&"</td><td>"&tx&"</td><td>"&dj&"</td><td>"&tu&"</td><td>"&sj&"</td><td>"&zt&"</td><td><a href=chgxx.asp?id="&id&"&opt=修改>修改</a> | <a href=CHGxx.ASP?id="&id&"&opt=删除>删除</a> </td></tr>"
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
<p align=center>轩辕-出家-鬼魂-武功管理</p>
<hr><br>
当武功类型显示为“轩”表明该项为轩辕武功，当为“佛”标明为出家人修炼的武功，当为“鬼”表明为鬼魂修炼的武功！
<table border=3 width=100% align=center>
<tr bgcolor=#FFFF00 align=center><td>ID</td><td>招式</td><td>所有者</td><td>武功类型</td><td>攻击</td><td>防御</td><td>速度</td><td>特效</td><td>等级</td><td>图片</td><td>时间</td><td>状态</td><td><a href="chgxx.asp?opt=新增&amp;id=-1">新增</a></td></tr>
<%=msg%>
</table>
</body>