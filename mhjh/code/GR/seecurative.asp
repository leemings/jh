<%
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 名称,体力,内力,攻击,防御,特效,数量 from 物品 where 属性='药品' and 所有者='"&username&"' and 数量>0 order by 价格"
rst.Open sqlstr,conn
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'><tr align=center bgcolor=ffffff><td>名称</td><td>体力</td><td>内力</td><td>攻击</td><td>防御</td><td>特效</td><td>数量</td></tr>"
do while not (rst.BOF or rst.EOF)
cname=rst("名称")
msg=msg&"<tr><td ><a href='usecur.asp?mg="&cname&"'>"&cname&"</td><td>"&rst("体力")&"</td><td>"&rst("内力")&"</td><td>"&rst("攻击")&"</td><td>"&rst("防御")&"</td><td>"&rst("特效")&"</td><td>"&rst("数量")&"</td></tr>"
rst.MoveNext
loop
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="../style.css"></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false marginwidth="5" marginheight="0">
<div align=center>
<b><font color="#000000" face="隶书" size="4">我的药品</font></b><hr>
<%=msg%>
</div>
</body>
