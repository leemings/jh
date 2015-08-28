<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 名称,数量 from 物品 where 属性='礼品' and 所有者='"&username&"' and 数量>0 order by 价格"
rst.Open sqlstr,conn
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'>"
do while not (rst.BOF or rst.EOF)
msg=msg&"<tr><td><a href=javascript:liping('"&rst("名称")&"');>"&rst("名称")&"</a></td><td width=15  bgcolor=ffff00>"&rst("数量")&"</td></tr>"
rst.MoveNext
loop
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="css3.css">
<script language=javascript>
function liping(ping){
location.href='liping.asp?mg='+ping+'&st='+parent.talkfrm.document.talkform.sendto.value;
}
</script>
</head>
<body background='bg1.gif' oncontextmenu=self.event.returnValue=false leftmargin="5" topmargin="0" marginwidth="5" marginheight="0">
<div align=center>
<font color=0000FF>使用礼品</font><br>
<hr noshade size="1" color=red>
<%=msg%>
</div>
</body>

