<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 门派,简介 from 门派 where 加入条件<>'False' order by 工资系数"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF)
corp=rst("门派")
msg=msg+"<tr><td bgcolor=ffff00 align=center><a href='joincorp.asp?mg="&corp&"' omouseover="&chr(34)&"window.status='加入"&corp&"';return true;"&chr(34)&"  onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&rst("简介")&">【"&corp&"】</a></td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="css2.css">
</head>
<body  oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5" background='bg1.gif'>
<div align=center><br><font color=008800><%=Application("yx8_mhjh_allonlinenum")+3%></font>人在线
<hr noshade size="1" color=red>
<font color=0000FF><b>加入门派</b></font><br>
<font color="#FF0000"><b>[</b>停留可看说明单击加入<b>]</b></font><br>
<table width=98% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
<%=msg%>
</table>
</div>
</body>
