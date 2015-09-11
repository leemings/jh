<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
cmid=request("id")
if not isnumeric(cmid) then
 cmid=1
end if
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open application("aqjh_usermdb")
Set rst= Server.CreateObject("ADODB.Recordset")
sqlstr="select * from 寄售 where id="&cmid&" and 数量>0"
rst.open sqlstr,connb
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('对不起，没有找到此类物品!'); window.close();</script>"
Response.End
end if
syz=rst("所有者")
name=rst("名称")
danjia=rst("价格")
sl=rst("数量")
rst.Close
set rst=nothing
connb.Close
set connb=nothing
%>
<head><title>自由集市</title><LINK href="../../css.css" rel=stylesheet></head>
<body background="../../Bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
<td bgcolor=#ffff00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("aqjh_chatroomname")%>自由集市</b></font></p>
<p align=center>购买物品：您需要带够现金，购买成功后，银两直接划出</p>
<div align=center>
<form action="gouok.asp?id=<%=cmid%>" method=post>
<table width='60%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
<tr><td width=36>物品</td><td><%=name%></td></tr>
<tr><td>标价</td><td><%=danjia%></td></tr>
<tr><td width=36>出售人</td><td><%=syz%></td></tr>
<tr><td>数量</td><td><input type=text name=sl value='<%=sl%>' maxlength=4 size=9></td></tr>
<tr>
<td colspan=2 align=center height="32">
<input type=submit value=" 购 买 "> <input type=reset value=" 重 置 "> <input type=button onclick="javascript:history.back()" value=" 返 回 "></td></tr>
</table>
</form>
</div></body>