<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usercorp="官府" and usergrade=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&"#"
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
set conn=server.CreateObject("adodb.connection")
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 姓名,门派,身份,等级,会员时间 from 用户 where 会员=true and 会员时间>"&nowtimetype&" order by 会员时间 desc",conn,1,2
if rst.EOF then
	msg="<tr align=center><td>现在没有人是会员</td></tr>"
else
	rst.PageSize=20
	if activepage>rst.PageCount then activepage=rst.PageCount
	rst.AbsolutePage=activepage
	for i=1 to rst.Pagesize
	if rst.EOF or rst.BOF then exit for
		msg=msg&"<tr><td><a href='#' onclick=javacript:document.form2.username.value='"&rst("姓名")&"'; onmouserover="&chr(34)&"window.status='选定管理对象为"&rst("姓名")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("姓名")&"</a></td><td>"&rst("门派")&"</td><td>"&rst("身份")&"</td><td>"&rst("等级")&"</td><td>"&formatdatetime(rst("会员时间"),1)&"</td></tr>"
		rst.MoveNext
	next
	msg=msg&"</table><table border=3 width='100%'><form action='showfell.asp' method=post id=form1 name=form1><tr align=center bgcolor=cccccc>"
		if activepage>1 then
	msg=msg&"<td><a href='showfell.asp?activepage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='showfell.asp?activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
	else
		msg=msg&"<td>第一页</td><td>前一页</td>"
	end if
	if activepage<rst.PageCount then
		msg=msg&"<td><a href='showfell.asp?activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='showfell.asp?activepage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
	else
		msg=msg&"<td>后一页</td><td>最后一页</td>"
	end if
	msg=msg&"<td>第<input type=text name=activepage size=4 value='"&activepage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></form>"
end if
%>
<html>
<head>
<link rel='stylesheet' href='../style.css'>'
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=4>会员管理</font></p><hr>
<table width=80% align=center border=3><tr bgcolor=ffff00 align=center><td>姓名</td><td>门派</td><td>身份</td><td>等级</td><td>会员到期时间</td></tr><%=msg%></table>
<form action=upfellow.asp method=post name=form2><table width=80% align=center border=3><tr><td>姓名</td><td><input type=text name='username' value='' maxlength=7 size=14></td><td>时间</td><td> <input type=text name='month' value='1' maxlength=2 size=2>个月</td><td colspan=2> <input type=submit value=' 增 加 ' name=submit> <input type=submit value=' 减 少 ' name=submit> <input type=reset value=' 重 置 '> </td></td></tr></table></form>
</body>
</html>