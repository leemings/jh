<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,名称,体力,内力,特效,价格 from 商店 where 属性='药品' order by 价格 "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
	rst.PageSize=10
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF><tr bgcolor=ffFF00 align=center><td width=100>名称</td><td>体力</td><td>内力</td><td>特效</td><td>价格</td></TR>"
	for j=1 to rst.Pagesize
		if rst.EOF or rst.BOF then
			exit for
		else	
			msg=msg&"<tr><td><a href='buy.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='购买"&rst("名称")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("名称")&"</td><td>"&rst("体力")&"</td><td>"&rst("内力")&"</td><td>"&rst("特效")&"</td><td>"&rst("价格")&"</td></tr>"
			rst.MoveNext
		end if
	next
	msg=msg+"</table><form action='curativeshop.asp' method=post><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center>"
	if acpage>1 then
		msg=msg&"<td><a href='curativeshop.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='curativeshop.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
	else
		msg=msg&"<td>第一页</td><td>前一页</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='curativeshop.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='curativeshop.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
	else
		msg=msg&"<td>后一页</td><td>最后一页</td>"
	end if
	msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页共"&rst.PageCount&"页<input type=submit value='GO'></td></tr></table></form>"
end if
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
%>
<head><LINK href="../chatroom/css.css" rel=stylesheet><title>药店</title></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("Ba_jxqy_systemname")%>药品店</b></font></p>
<div align=center><%=msg%></div>
<p align=center>
  <input type="button" value=" 返 回 " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p>

</body>