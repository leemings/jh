<%
Response.Expires=-1
bgcolor=Application("yx8_mhjh_backgroundcolor")
cmid=Request.QueryString("id")
if not isnumeric(cmid) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,名称,价格 from 物品 where id="&cmid&" and 所有者='"&username&"' and 数量>0 and 寄售=False "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=028"
cmname=rst("名称")
cmprice=rst("价格")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>自由集市</title><LINK href="../style4.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>'  oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
<td bgcolor=#005b00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
<td><A href="chgprice.asp" onmouseover="window.strtus='变更已寄售物品的价格';return true;" onmouseout="window.status='';return true">改价</a></td>
</tr></table>
<p align=center><b><font color="#CC0000" face="方正舒体" size="6">自由集市</font></b></p>
<div align=center>
<form action="sell.asp?id=<%=cmid%>" method=post>
<table width='60%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
<tr><td width=36>物品</td><td><%=cmname%></td></tr>
<tr><td>标价</td><td><input type=text name=price value='<%=cmprice%>' maxlength=9 size=9></td></tr>
<tr>
<td colspan=2 align=center height="32">
<input type=submit value=" 寄 售 "> <input type=reset value=" 重 置 "> <input type=button onclick="javascript:history.back()" value=" 返 回 "></td></tr>
</table>
</form>
</div>
<p align=center>
　
</p>
</body>
