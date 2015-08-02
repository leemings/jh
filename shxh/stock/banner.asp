<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><table align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 证交公告 order by 时间 desc",conn,1,3
if rst.EOF or rst.BOF then
	msg=msg&"<tr align=venter>没有公告</tr></table>"
else
	rst.PageSize=1
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	msg=msg&"<tr align=center bgcolor=FFFF00><td><b><font face=宋体 size=2 color=red>"&rst("标题")&"</font></b></td></tr><tr><td > 【<font color=ff0000>证</font>】字第"&rst("id")&"号</td></tr><tr><td>"&rst("内容")&"</td></tr><tr><td align=right>"&formatdatetime(rst("时间"),1)&"</td></tr></table><form action='banner.asp' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center>"
	if acpage>1 then
		msg=msg&"<td><a href='banner.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='banner.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
	else
		msg=msg&"<td>第一页</td><td>前一页</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='banner.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='banner.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
	else
		msg=msg&"<td>后一页</td><td>最后一页</td>"
	end if
	msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg&"</body>"
%>