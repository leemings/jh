<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 名称,特效,数量 from 物品 where 属性='暗器' and 所有者='"&username&"' and 数量>0 order by 价格"
rst.Open sqlstr,conn
msg="<table width=95% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF><tr></tr>"
do while not (rst.BOF or rst.EOF)
msg=msg&"<tr><td><a href=javascript:parent.talkfrm.settalk('//投掷','"&rst("名称")&"') target='talkfrm' onmouseover=""window.status='使用暗器';return true;"" onmouseout=""window.status='';return true;"" title="&chr(34)&rst("特效")&chr(34)&">"&rst("名称")&"</a></td><td width=15  bgcolor=ffff00>"&rst("数量")&"</td></tr>"
rst.MoveNext
loop
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="css3.css"></head>
<body  oncontextmenu=self.event.returnValue=false leftmargin="4" marginwidth="4" background='bg1.gif'>
<div align=center><br>
<font color=008800><%=Application("yx8_mhjh_allonlinenum")%></font>人在线
<hr noshade size="1" color=red>
<font color=0000FF>拥有暗器</font><br>
<%=msg%>
</div>
<br>&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.open('../js/armorshop.asp','reg','width=580,height=550,scrollbars=yes')"><font color="#F71004">去武器暗器店</font></a> 
</body>
