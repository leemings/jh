<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 招式,基本攻击 from 攻击 where 姓名='"&username&"' order by 基本攻击"
rst.Open sqlstr,conn
msgarr=";"
do while not (rst.EOF or rst.BOF)
attackname=rst("招式")
msgarr=msgarr&attackname&";"
msg=msg+"<tr><td bgcolor=FEE3AB><a href=javascript:parent.settalk('//攻击','"&attackname&"'); target='talkfrm' onmouseover=""window.status='使用武功攻击他人';return true;"" onmouseout=""window.status='';return true;"" title='使用武功攻击他人'><font color=FF0000 valign='middle'>"&attackname&"</font></a></td><td>"&rst("基本攻击")&"</td></tr>"
rst.movenext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="<tr valign=middle height='80%'><td colspan=2>小孩子，对不起了，你都没有学会任何招式，又怎样攻击呢</td></tr>"
%>
<head><link rel="stylesheet" href="css.css"></head>
<body  oncontextmenu=self.event.returnValue=false leftmargin="5" topmargin="2" marginwidth="5" marginheight="2" background='bg1.gif'>
<div align=center><br>
<font color=FF0000><font color=008800><%=Application("yx8_mhjh_allonlinenum")%></font>人在线
<hr noshade size="1" color=darkred>
<font color=0000FF>招式攻击</font><br>
<table width='99%' border=1 cellspacing="1" cellpadding="1" align="right" bordercolor="#FF9933">
<tr align=center>
<td height="19" width="65%">招式</td>
<td height="19" width="36%">攻击</td>
</tr>
<%=msg%>
</table>
</div>
</body>
