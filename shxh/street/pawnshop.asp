<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 物品 where 数量>0 and 所有者='"&username&"' order by 价格 "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
	rst.PageSize=10
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg="<table border=3 width='100%' ><tr bgcolor=ffFF00 align=center><td width=40>类型</td><td width=100>名称</td><td>参考价格</td><td width=36>寄售</td></TR>"
	for j=1 to rst.Pagesize
		if rst.EOF or rst.BOF then
			exit for
		else	
			msg=msg&"<tr><td>"&rst("属性")&"</td><td>"&rst("名称")&"</td><td align=right>"&rst("价格")&"</td><td><a href='pawn.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='当了"&rst("名称")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">当了</a></td></tr>"
			rst.MoveNext
		end if
	next
	msg=msg+"</tr></table><form action='pawnshop.asp' method=post id=form1 name=form1><table border=1 width=100% bgcolor=ffff00><tr align=center>"
	if acpage>1 then
		msg=msg&"<td><a href='pawnshop.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='pawnshop.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
	else
		msg=msg&"<td>第一页</td><td>前一页</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='pawnshop.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='pawnshop.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
	else
		msg=msg&"<td>后一页</td><td>最后一页</td>"
	end if
	msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
if msg="" then msg="你没有东西可供典当！"
%>
<head><title>当铺</title><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000FF size=6 face='幼圆'><%=Application("Ba_jxqy_systemname")%>当铺</font></p>
<div align=center><%=msg%></div>
<p align=center>
  <input type="button" value=" 返 回 " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p></body>