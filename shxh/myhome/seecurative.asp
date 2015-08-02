<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 名称,体力,内力,特效,数量 from 物品 where 属性='药品' and 所有者='"&username&"' and 数量>0 order by 价格"
rst.Open sqlstr,conn
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'><tr align=center bgcolor=ffff00><td>名称</td><td>体力</td><td>内力</td><td>特效</td><td>数量</td></tr>"
do while not (rst.BOF or rst.EOF)
	cname=rst("名称")
	msg=msg&"<tr><td ><a href='usecur.asp?mg="&cname&"'>"&cname&"</td><td>"&rst("体力")&"</td><td>"&rst("内力")&"</td><td>"&rst("特效")&"</td><td>"&rst("数量")&"</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table>"
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false marginwidth="5" marginheight="0">
<div align=center>
<font color=0000FF size=4>拥有药品</font><hr>
<%=msg%>
<input type=button value='返回' onclick=javascript:location.href='myhome.asp';>
</div>
</body>
