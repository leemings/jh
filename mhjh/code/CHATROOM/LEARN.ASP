<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
mycorp=session("yx8_mhjh_usercorp")
username=session("yx8_mhjh_username")
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 招式,等级 from 攻击 where 姓名='"&username&"' order by 等级 desc"
rst.Open sqlstr,conn
msgarr=";"
do while not (rst.EOF or rst.BOF)
attackname=rst("招式")
msgarr=msgarr&attackname&";"
msg=msg+"<tr><td><a href='learn2.asp?mg="&attackname&"' onmouseover=""window.status='升级业已学会的功夫';return true;"" onmouseout=""window.status='';return true;"" title='升级业已学会的功夫'>"&attackname&"</a></td><td>"&rst("等级")&"级</td></tr>"
rst.movenext
loop
rst.Close
sqlstr="select 招式 from 招式 where 门派='"&mycorp&"'"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF)
attackname=rst("招式")
if instr(msgarr,";"&attackname&";")=0 then msg=msg+"<tr><td><a href='learn2.asp?mg="&attackname&"' onmouseover=""window.status='修习新的功夫';return true;"" onmouseout=""window.status='';return true;"" title='修习新的功夫'><font color=FF0000>"&attackname&"</font></a></td><td>0级</td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="<tr valign=middle height='80%'><td colspan>对不起了，没有你合适修炼的武功</td></tr>"
%>
<head><link rel="stylesheet" href="css.css"></head>
<body oncontextmenu=self.event.returnValue=false background='bg1.gif'>
<div align=center><br>
<font color=008800><%=Application("yx8_mhjh_allonlinenum")%></font>人在线<hr>
<font color=0000FF>修习武功</font><br>
<table width="80%" border="0" >
<%=msg%>
</table>
</div>
</body>
