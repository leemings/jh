<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 门派,简介,体力 from 门派 where 加入条件<>'False'"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF)
corp=rst("门派")
tl=rst("体力")
msg=msg+"<tr><td><a href='javascript:parent.chgsendto("&chr(34)&corp&chr(34)&");' target='talkfrm' onmouseover=""window.status='攻打该门派的要塞';return true;"" onmouseout=""window.status='';return true;"" title='攻打该门派的要塞'><font color=#CC0000>"&corp&"</font></a></td><td>"&tl&"</td></tr>"
rst.movenext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="<tr valign=middle height='80%'><td colspan=2>你眼睛有毛病？对不起！没有这样的门派让你攻击！是不是想称霸江湖想疯了？去买点抗疯狂的药品吃吧！蠢货！</td></tr>"
%>
<head><link rel="stylesheet" href="css2.css">
</head>
<body  oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5" background='bg1.gif'>
<div align=center><br><font color=008800><%=Application("yx8_mhjh_allonlinenum")%></font>人在线
<hr noshade size="1" color=red>
<font color=0000FF><b>攻打门派要塞</b></font><br>
<font color="#FF0000"><b>[</b>点门派名字，再发言，就可以攻击要塞，一统江湖的大业将在这里实现！<b>]</b></font><br>
<table width=98% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
<%=msg%>
</table>
</div>
</body>
