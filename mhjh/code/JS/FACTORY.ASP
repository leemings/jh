<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来介绍所。谢谢合作');location.href = '';</script>"
Response.End 
end if
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
msg="<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel=stylesheet href='../style.css'></head><body background='../chatroom/bg1.gif' oncontextmenu='self.event.returnValue=false;'><center><font color=red>10级以下新手补给地带，注意:您的体力上限是:等级X100万！</font><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center ><tr align=center bgcolor=ffffff><td width=70>工作</td><td>说明</td><td width=30>报酬</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from 打工 order by 银两",conn,1,3
if not (rst.EOF or rst.BOF) then
rst.PageSize=12
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
for i=1 to rst.PageSize
if rst.EOF then exit for
		msg=msg&"<tr><td><input type='button' value="&rst("工作")&" onClick=javascript:location.href='work.asp?id="&rst("id")&"' name='button'></a></td><td>"&rst("说明")&"</td><td align=right>"&rst("银两")&"</td></tr>"
		rst.Movenext
next
msg=msg&"</table><form action='factory.asp' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffffff><tr align=center>"
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
msg=msg&"<p align=center></p></body>"
Response.Write msg
%>
























