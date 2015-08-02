<%
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
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
<head><link rel="stylesheet" href="style1.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="4" marginwidth="4">
<div align=center><font size="4" color="#CC0000" face="幼圆"><b><%=Application("Ba_jxqy_systemname"&chatroomsn)%></b></font><br>
  <font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>人在线
  <hr noshade size="1" color=red>
<font color=0000FF>拥有暗器</font><br>
<%=msg%>
</div>
</body>
