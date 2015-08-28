<!--#include file="data.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
corp=Request.QueryString("corp")
corpa=Request.QueryString("corpa")
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
bgcolor=Application("yx8_mhjh_backgroundcolor")
msg="<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel=stylesheet href='../../style.css'></head><body background='../../chatroom/bg1.gif' oncontextmenu='self.event.returnValue=false;'><center><font color=red>这里是点歌台，每点一首歌，银两减20，体力减10<br>不能播放歌曲的请下载<a href='http://www.21ex.com/rp7setup.exe'>Replay网络播放软件</a></font><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center ><tr align=center bgcolor=ffffff><td >歌曲名字</td><td>歌手</td><td>人气</td><td>试听</td></tr>"
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from 歌曲 where "&corp&"='"&corpa&"' order by 点击 desc",conn,1,3
if not (rst.EOF or rst.BOF) then
rst.PageSize=30
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
for i=1 to rst.PageSize
if rst.EOF then exit for
		msg=msg&"<tr><td>"&rst("歌名")&"</td><td>"&rst("歌手")&"</td><td align=center>"&rst("点击")&"</td><td align=center><input onClick=javascript:window.open('song.asp?id="&rst("id")&"','yhjk','left=0,top=0,width=275,height=116,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no') type=button value='试听' name='button'></td></tr>"
		rst.Movenext
next
msg=msg&"</table><form action='show.asp' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffffff><tr align=center>"
if acpage>1 then
msg=msg&"<td><a href='show.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='show.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
msg=msg&"<td><a href='show.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='show.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
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
















