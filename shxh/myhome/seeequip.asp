<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='80%'><tr bgcolor=ffff00 align=center><td>属性</td><td>名称</td><td>攻击</td><td>防御</td><td>特效</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 属性,名称,攻击,防御,特效 from 物品 where 所有者='"&username&"' and 装备=true order by 属性",conn
do while not rst.EOF
	msg=msg&"<tr><td width=35 bgcolor=#FFff00><font color=red>"&rst("属性")&"</font></td><td ><a href='discharge.asp?cmtype="&rst("属性")&"' onmouseover=""window.status='卸下';return true;"" onmouseout=""window.status='';return true;"">"&rst("名称")&"</a></td><td>"&rst("攻击")&"</td><td>"&rst("防御")&"</td><td>"&rst("特效")&"</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table><font color='#006600'>【拥有装备】</font><table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='80%'><tr bgcolor=ffff00 align=center><td>属性</td><td>名称(X数量)</td><td>攻击</td><td>防御</td><td>特效</td></tr>"
rst.Close
sqlstr="select 名称,属性,数量,攻击,防御,特效 from 物品 where 属性 in('头盔','盔甲','武器','防具','饰品') and 所有者='"&username&"' and 数量>0 order by 价格"
rst.Open sqlstr,conn
do while not (rst.BOF or rst.EOF)
	msg=msg&"<tr><td width=35 bgcolor=#FFCC00><font color=red>"&rst("属性")&"</font></td><td><a href='equip.asp?commodityname="&rst("名称")&"' onmouseover=""window.status='装备';return true;"" onmouseout=""window.status='';return true;"">"&rst("名称")&"</a>(x"&rst("数量")&")</td><td>"&rst("攻击")&"</td><td>"&rst("防御")&"</td><td>"&rst("特效")&"</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table>"
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" topmargin="0" marginwidth="5" marginheight="0">
<div align=center><font size="4" color="#990000"><b><font face="幼圆"><%=Application("Ba_jxqy_systemname"&chatroomsn)%></font></b></font><br>
   <hr noshade size="1">
  <font color="#006600">【查看装备】</font><br>
<%=msg%>
<input type=button value='返回' onclick=javascript:location.href='myhome.asp';>
</div>
</body>
