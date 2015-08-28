<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 属性,名称 from 物品 where 所有者='"&username&"' and 装备=true order by 属性",conn
do while not rst.EOF
msg=msg&"<tr><td width=35 bgcolor=#FFff00><font color=red>"&rst("属性")&"</font></td><td ><a href='discharge.asp?cmtype="&rst("属性")&"' onmouseover=""window.status='卸下';return true;"" onmouseout=""window.status='';return true;"">"&rst("名称")&"</a></td></tr>"
rst.MoveNext
loop
msg=msg&"</table><font color='#006600'>【拥有装备】</font><table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'>"
rst.Close
sqlstr="select 名称,属性,数量 from 物品 where 属性 in('头盔','盔甲','武器','防具','饰品') and 所有者='"&username&"' and 数量>0 order by 价格"
rst.Open sqlstr,conn
do while not (rst.BOF or rst.EOF)
msg=msg&"<tr><td width=35 bgcolor=#FFCC00><font color=red>"&rst("属性")&"</font></td><td><a href='equip.asp?commodityname="&rst("名称")&"' onmouseover=""window.status='装备';return true;"" onmouseout=""window.status='';return true;"">"&rst("名称")&"</a>(x"&rst("数量")&")</td></tr>"
rst.MoveNext
loop
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="css3.css"></head>
<body  oncontextmenu=self.event.returnValue=false leftmargin="5" topmargin="0" marginwidth="5" marginheight="0" background='bg1.gif'>
<div align=center><br>
<font color="#006600">【查看装备】</font><br>
<%=msg%>
</div>
<br>&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.open('../js/armorshop.asp','reg','width=580,height=550','scrollbars=yes')"><font color="#F71004">去武器暗器店</font></a> 
</body>
