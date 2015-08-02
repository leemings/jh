<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 名称,特效,数量 from 物品 where 属性='药品' and 所有者='"&username&"' and 数量>0 order by 价格"
rst.Open sqlstr,conn
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'>"
do while not (rst.BOF or rst.EOF)
	msg=msg&"<tr><td><a href=javascript:usercur('"&rst("名称")&"'); onmouseover=""window.status='使用药品';return true;"" onmouseout=""window.status='';return true;"" title="&chr(34)&rst("特效")&chr(34)&">"&rst("名称")&"</a></td><td width=15  bgcolor=ffff00>"&rst("数量")&"</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table>"
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="style1.css">
<script language=javascript>
function usercur(cur){
	location.href='usecur.asp?mg='+cur+'&st='+parent.talkfrm.document.talkform.sendto.value;
}
</script>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" topmargin="0" marginwidth="5" marginheight="0">
<div align=center>
<font color=0000FF size=5>使用药品</font><br>
<hr>
<%=msg%>
</div>
</body>
