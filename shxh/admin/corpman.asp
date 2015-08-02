<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
corp=Request.QueryString("corp")
activepage=Request.QueryString("activepage")
if instr(corp,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject ("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"' and 身份='掌门' and 门派='"&corp&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=046"
rst.Close
rst.Open "select * from 用户 where 门派='"&corp&"'",conn,1,2
rst.PageSize=20
if activepage>rst.PageCount then activepage=rst.PageCount
rst.AbsolutePage=activepage
i=0
msg="<tr align=center bgcolor=00ff00><td>姓名</td><td>等级</td><td>身份</td><td>体力</td><td>内力</td><td>攻击</td><td>防御</td></tr>"
do while i<rst.PageSize
	if rst.EOF or rst.BOF then exit do 
	if rst("身份")="掌门" then
		msg=msg&"<tr><td>"&rst("姓名")&"</td>"
	else
		msg=msg&"<tr><td><a href='#' onclick="&chr(34)&"javascript:document.form1.st.value='"&rst("姓名")&"';document.form1.degree.value='"&rst("身份")&"';"&chr(34)&">"&rst("姓名")&"</a></td>"
	end if
	msg=msg&"<td>"&rst("等级")&"</td><td>"&rst("身份")&"</td><td>"&rst("体力")&"</td><td>"&rst("内力")&"</td><td>"&rst("攻击")&"</td><td>"&rst("防御")&"</td></tr>"
	rst.MoveNext
	i=i+1
loop
msg=msg&"</table><form id=form1 name=form1 action=corpman2.asp method=post><table border=0 width=100% align=center><tr><td>门人<input type=text name='st' value='' maxlength=7 size=14></td><td>身份<input type=text name='degree' value='无' mxalength=7 size=14></td><td><input type=submit name=submit value='册封'> <input type=submit name=submit value='逐出门墙'></td></tr></table></form>"
if activepage<>1 then 
	msg=msg&"<table border=3  width=100% align=center><tr><td><a href='corpman.asp?corp="&corp&"&activepage=1'>第一页</a></td><td><a href='corpman.asp?corp="&corp&"&activepage="&activepage-1&"'>前一页</a></td>"
else
	msg=msg&"<table border=3  width=100% align=center><tr><td>第一页</td><td>前一页</td>"
end if
if activepage=rst.PageCount then
	msg=msg&"<td>后一页</td><td>最后一页</td></tr>"
else
	msg=msg&"<td><a href='corpman.asp?corp="&corp&"&activepage="&activepage+1&"'>后一页</a></td><td><a href='corpman.asp?activepage="&rst.PageCount&"&corp="&corp&"'>最后一页</a></td></tr>"
end if
set rst=nothing
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../chatroom/style1.css'></head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"' topmargic=100><p align=center><font color=ff0000 size=5>门户管理</font><hr></p><table width=80% align=center border=3>"&msg&"</table></body>"
%>