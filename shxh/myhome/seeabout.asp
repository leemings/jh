<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 姓名,性别,门派,身份,配偶,精力,积分,等级,银两,体力,内力,攻击,防御,资质,道德,会员,会员时间,protect from 用户 where 姓名='"&username&"'",conn
for i=0 to 8
	msg1=msg1&"<td>"&rst.Fields(i).Value&"</td>"
	msg2=msg2&"<td>"&rst.Fields(i+9).Value&"</td>"
next
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<LINK href="../chatroom/css.css" rel=stylesheet>
<div align=center>
<font color=0000FF>基本状态</font><hr>
  <table width=95% bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF border=3>
    <tr bgcolor=ffff00 align=center><td>姓名</td><td>性别</td><td>门派</td><td>身份</td><td>配偶</td><td>精力</td><td>积分</td><td>等级</td><td>银两</td></tr>
    <tr><%=msg1%></tr>
    <tr bgcolor=ffff00 align=center><td>体力</td><td>内力</td><td>攻击</td><td>防御</td><td>资质</td><td>道德</td><td>会员</td><td>会员时间</td><td>protect</td></tr>
    <tr><%=msg2%></tr>
  </table>
 <input type=button value='返回' onclick=javascript:location.href='myhome.asp'; id=button1 name=button1>
</div>
</body>
