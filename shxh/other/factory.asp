<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel=stylesheet href='../chatroom/css.css'></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;'><center><font color=red>工坊不接受江湖高手打工[切记平衡发展各属性]</font><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center ><tr align=center bgcolor=FFFF00><td width=70>工作</td><td>说明</td><td width=30>报酬</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from 打工 order by 银两",conn,1,3
if not (rst.EOF or rst.BOF) then
	rst.PageSize=12
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	for i=1 to rst.PageSize
		if rst.EOF then exit for
		msg=msg&"<tr><td><a href='work.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='参加"&rst("工作")&"工作';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"参加"&rst("工作")&"工作"&chr(34)&">"&rst("工作")&"</a></td><td>"&rst("说明")&"</td><td align=right>"&rst("银两")&"</td></tr>"
		rst.Movenext
	next
	msg=msg&"</table><form action='factory.asp' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center>"
	if acpage>1 then
		msg=msg&"<td><a href='factory.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='factory.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
	else
		msg=msg&"<td>第一页</td><td>前一页</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='factory.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='factory.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
	else
		msg=msg&"<td>后一页</td><td>最后一页</td>"
	end if
	msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
else
	msg=msg&"</table>"	
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"<p align=center><input type=button value=' 返 回 ' onClick=javascript:location.href='../street/street.asp';> <input type=button value=' 关 闭 ' onClick=javascript:window.close();></p></body>"
Response.Write msg
%>