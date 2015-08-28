<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
guojia=trim(Request.QueryString("guojia"))
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
bgcolor=Application("yx8_mhjh_backgroundcolor")
msg="<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel=stylesheet href='css.css'></head><body  bgcolor='"&bgcolor&"' background=bg1.gif oncontextmenu='self.event.returnValue=false;' style='CURSOR: url(../images/banana.ani)'><center><img border=0 src='../chatroom/Image/yys.gif' width=350 height=22><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center ><tr align=center bgcolor=EEF7F7><td >姓名</td><td>国家职位</td><td>门派</td><td>任务</td><td>等级</td><td>积分</td><td>银两</td></tr>"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select 姓名,职位,门派,任务,等级,积分,银两 from 用户 where 国家='"&guojia&"' order by 积分 desc",conn,1,3
if not (rst.EOF or rst.BOF) then
rst.PageSize=12
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
for i=1 to rst.PageSize
if rst.EOF then exit for
		msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("职位")&"</td><td>"&rst("门派")&"</td><td>"&rst("任务")&"</td><td>"&rst("等级")&"</td><td align=center>"&rst("积分")&"</td><td align=center>"&rst("银两")&"</td></tr>"
		rst.Movenext
next
msg=msg&"</table><form action='guojia2.asp' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=#EEF7F7><tr align=center>"
if acpage>1 then
msg=msg&"<td><a href='guojia2.asp?acpage=1&guojia="&guojia&"' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='guojia2.asp?acpage="&acpage-1&"&guojia="&guojia&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
msg=msg&"<td><a href='guojia2.asp?acpage="&acpage+1&"&guojia="&guojia&"'  onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='guojia2.asp?acpage="&rst.PageCount&"&guojia="&guojia&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
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