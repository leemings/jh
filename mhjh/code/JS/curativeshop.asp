<%
Response.Expires=-1
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,名称,体力,内力,特效,价格 from 商店 where 属性='药品' order by 价格 "
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
rst.PageSize=10
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
msg="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=ffffff><tr bgcolor=ffffff align=center><td width=100>名称</td><td>体力</td><td>内力</td><td>特效</td><td>价格</td><td>数量</td><td>操作</td></TR>"
for j=1 to rst.Pagesize
if rst.EOF or rst.BOF then
exit for
else
msg=msg&"<tr><form method=POST action='buy.asp'><td><input type=hidden checked value='"&rst("id")&"' name='id'>"&rst("名称")&"</td><td>"&rst("体力")&"</td><td>"&rst("内力")&"</td><td>"&rst("特效")&"</td><td>"&rst("价格")&"</td><td><select name='sl'><option value='1' selected>1</option><option value='2'>2</option><option value='3'>3</option><option value='4'>4</option><option value='5'>5</option><option value='6'>6</option><option value='7'>7</option><option value='8'>8</option><option value='9'>9</option></select></td><td><input type='SUBMIT' name='Submit' value='购买'></td></form></tr>"
rst.MoveNext
end if
next
msg=msg+"</table><form action='curativeshop.asp' method=post><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=ffffff bgcolor=ffffff><tr align=center>"
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
<head><LINK href="../style.css" rel=stylesheet><title>药店</title></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align=center><b><font face="方正舒体" size="3">山草药店</font></b></p>
<div align=center><%=msg%></div>
<p align=center>
<input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p>


</body>

