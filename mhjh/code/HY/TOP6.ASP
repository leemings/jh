<HTML><HEAD><TITLE>二十大美女</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 4.0" name=GENERATOR><style type="text/css">.lh22px {
FONT-SIZE: 12px; LINE-HEIGHT: 22px
}
BODY {
FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
TD {
FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
A:visited {
COLOR: #0000a0; TEXT-DECORATION: none
}
A:link {
COLOR: #0000a0; TEXT-DECORATION: none
}
A:hover {
COLOR: #ff0000; TEXT-DECORATION: underline
}
.haody A:visited {
COLOR: #e9edfe; TEXT-DECORATION: none
}
.haody A:link {
COLOR: #e9edfe; TEXT-DECORATION: none
}
.haody A:hover {
COLOR: #ffffff; TEXT-DECORATION: underline
}
</style><%
const MaxPerPage=20
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>
</HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=534>
<TBODY>
<TR>
<TD background=jh.files/history_table_bg.gif height=158 vAlign=top width="568">
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain
width=560>
<TBODY>
<TR>
<TD vAlign=top>
<p align="center"><font color="#FFFF00" size="4" face="隶书"><b>美女排行榜<BR>
</b></font>
<TABLE align=center background=jh.files/table_bg.gif border=0
cellPadding=0 cellSpacing=0 class=p9 width="100%">
<TBODY>
<TR>
<TD align=middle vAlign=center>
<P class=p9><br>
　　江湖中的二十大美女，呵呵，她们可都是热门人物。每天都有大侠们送她们鲜花，首饰，怪不得是越来越美了。如果你想在剑侠里成家立业的话，呵呵，还不向这二十大美女发起进攻。祝你好运！！！<br>
<%
dim sql
dim rs
dim filename
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select * from 用户 where  性别='女' order by 美貌 desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'>没有可排行的对象 </p>"
else
%>
</P>
<table border="0" cellspacing="1" width="562" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0">
<tr bgcolor="#336633">
<td align="center" height="18" width="93" bgcolor="#CC6666"><font  color="#FFFFFF">芳名</font></td>
<td align="center" height="18" width="93" bgcolor="#CC6666"><font  color="#FFFFFF">门派</font></td>
<td align="center" height="18" width="93" bgcolor="#CC6666"><font  color="#FFFFFF">身份</font></td>
<td align="center" height="18" width="93" bgcolor="#CC6666"><font  color="#FFFFFF">美貌</font></td>
<td align="center" height="18" bgcolor="#CC6666" width="94"><font  color="#FFFFFF">个性</font></td>
<td align="center" height="18" bgcolor="#CC6666" width="94"><font  color="#FFFFFF">伴侣</font></td>
</tr>
<%do while not rs.eof%>
<tr>
<td align="center" bgcolor="#FF9F71" width="93"><%=rs("姓名")%>
</td>
<td align="center" bgcolor="#FF9966" width="93"><%=rs("门派")%>
</td>
<td align="center" bgcolor="#FF9966" width="93"><%=rs("身份")%>
</td>
<td align="center" bgcolor="#FF9966" width="93"><font color=ff0000><%=rs("美貌")%></font>
</td>
<td align="center" bgcolor="#FF9966" width="94"><%=rs("个性")%>　 </td>
<td align="center" bgcolor="#FF9966" width="94"><%=rs("配偶")%> </td>
<%
rs.movenext
filename=filename+1
if filename>19 then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.Close             
set conn=nothing  
%>

</table>
<P align=center>&nbsp;</P></TD></TR></TBODY></TABLE>
</TD></TR></TBODY></TABLE></TD></TR>
</TBODY></TABLE></BODY></HTML>
