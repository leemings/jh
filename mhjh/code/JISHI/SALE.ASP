<%
Response.Expires=-1
bgcolor=Application("yx8_mhjh_backgroundcolor")
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 物品 where 所有者='"&username&"' and 数量>0 and 寄售=False order by 价格 "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
rst.PageSize=10
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
msg="<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FF66cc align='center'><tr bgcolor=005b00 align=center><td width=40>类型</td><td width=140>名称</td><td>参考价格</td><td width=36>寄售</td></TR>"
for j=1 to rst.Pagesize
if rst.EOF or rst.BOF then
exit for
else
msg=msg&"<tr><td>"&rst("属性")&"</td><td>"&rst("名称")&"</td><td>"&rst("价格")&"</td><td><a href='vend.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='寄售"&rst("名称")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">寄售</a></td></tr>"
rst.MoveNext
end if
next
msg=msg+"</tr></table><form action='sale.asp' method=post id=form1 name=form1><table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align='center' bgcolor=005b00><tr align=center>"
if acpage>1 then
msg=msg&"<td><a href='sale.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='sale.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
msg=msg&"<td><a href='sale.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='sale.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="你没有东西可供寄售！"
%>
<head><title>自由集市</title><LINK href="../style4.css" rel=stylesheet><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor='<%=bgcolor%>' background="bg1.gif"  oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
<td bgcolor=#005b00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
<td><A href="chgprice.asp" onmouseover="window.strtus='变更已寄售物品的价格';return true;" onmouseout="window.status='';return true">改价</a></td>
</tr></table>
<p align=center><b><font color="#FFFF00" size="5" face="方正舒体">自由集市</font></b></p>
<div align=center><%=msg%></div>
<p align=center>
&nbsp;
<input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p>
</body>
