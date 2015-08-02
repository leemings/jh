<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if not(session("Ba_jxqy_usercorp")="官府" and Session("Ba_jxqy_usergrade")>=Application("Ba_jxqy_exaltgraderight")) then Response.Redirect "../error.asp?id=046"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center><font color=0000FF size=4>官府人员管理</font></p><hr><table width=80% align=center border=3><tr align=center bgcolor=FFFF00><td>姓名</td><td>级别</td></tr>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "select * from 用户 where 门派='官府'",conn
do while not (rst.EOF or rst.BOF)
	Response.Write "<tr><td><a href='#' onclick="&chr(34)&"javascript:document.form1.uname.value='"&rst("姓名")&"';"&chr(34)&">"&rst("姓名")&"</td><td>"&rst("等级")&"</td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write "<tr align=center><td colspan=2><form action=upgov.asp method=post name=form1>姓名:<input type=text size=14 maxlength=14 name='uname'> <input type=submit name='submit1' value=' 升 级 '> <input type=submit name='submit1' value=' 降 级 '>  <input type='submit' name='submit1' value=' 开 除 '> <input type=submit name='submit1' value=' 聘 用 '></form></td></tr></table></body>"
%>