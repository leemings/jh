<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
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
sqlstr="select 时间,死者,凶手,死因 from 英烈堂 order by 时间 desc"
rst.open sqlstr,conn,1,2
if  not (rst.EOF or rst.BOF) then
rst.PageSize=20
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
msg="<table border=1 width='100%' cellpadding='0' cellspacing='0' bordercolorlight=000000 bordercolordark=FFFFFF align=center><tr bgcolor=ffffff align=center><td>时间</td><td>死者</td><td>凶手</td><td>死因</td></TR>"
for j=1 to rst.Pagesize
if rst.EOF or rst.BOF then
exit for
else
msg=msg&"<tr><td>"&rst("时间")&"</td><td>"&rst("死者")&"</td><td>"&rst("凶手")&"</td><td>"&rst("死因")&"</td></tr>"
rst.MoveNext
end if
next
msg=msg+"</tr></table><form action='ma.asp' method=post id=form1 name=form1><table border=1 width=100% cellpadding='0' cellspacing='0' bordercolorlight=000000 bordercolordark=FFFFFF align=center><tr align=center bgcolor=ffffff>"
if acpage>1 then
msg=msg&"<td><a href='ma.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='ma.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
msg=msg&"<td><a href='ma.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='ma.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="目前没有人死亡！"
%>
<head><title>江湖命案</title><link rel="stylesheet" href="../style.css"></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align=center><b><font face="隶书" size="3">江湖命案</font></b></p>
<div align=center><%=msg%></div>
</body>
