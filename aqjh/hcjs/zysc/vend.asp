<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
lx=request("lx")
name=request("name")
sl=request("sl")
danjia=request("danjia")
select case lx
case "物品"
dang="w6"
case "配药"
dang="w8"
case "魔宝"
dang="w9"
case "法器"
dang="w10"
case "装备"
dang="w3"
case else
Response.Write "<script language=javascript>alert('对不起，类型出错,想作弊吗?'); window.close();</script>"
Response.End
end select
if name="" then
Response.Write "<script language=javascript>alert('对不起，物品名出错!'); window.close();</script>"
Response.End
end if
if not isnumeric(sl) or not isnumeric(danjia) then
Response.Write "<script language=javascript>alert('对不起，数量或单价出错!'); window.close();</script>"
Response.End
end if
%>
<head><title>自由集市</title><LINK href="../../css.css" rel=stylesheet></head>
<body background="../../bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp">购买</a></td>
<td bgcolor=#ffff00><A href="sale.asp">寄售</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("aqjh_chatroomname")%>自由集市</b></font></p>
<p align=center>注：商店可以买到的东西禁止出售</p>
<div align=center>
<form action="sell.asp" method=post>
<input name="lx" type="hidden" value='<%=lx%>'>
<input name="name" type="hidden" value='<%=name%>'>
<table width='60%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center'>
<tr><td width=36>类型</td><td><%=lx%>类</td></tr>
<tr><td width=36>物品</td><td><%=name%></td></tr>
<tr><td>单价</td><td><input type=text name=danjia value='<%=danjia%>' maxlength=9 size=9></td></tr>
<tr><td>数量</td><td><input type=text name=sl value='<%=sl%>' maxlength=4 size=9></td></tr>
<tr>
<td colspan=2 align=center height="32">
<input type=submit value=" 寄 售 "> <input type=reset value=" 重 置 "> <input type=button onclick="javascript:history.back()" value=" 返 回 "></td></tr>
</table>
</form>
</div></body>