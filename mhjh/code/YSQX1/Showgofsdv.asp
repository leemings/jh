<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center><font color=0000FF size=4>�ٸ���Ա����</font></p><hr><table width=80% align=center border=3><tr align=center bgcolor=FFFF00><td>����</td><td>����</td></tr>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("yx8_mhjh_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "select * from �û� where ����='�ٸ�'",conn
do while not (rst.EOF or rst.BOF)
Response.Write "<tr><td><a href='#' onclick="&chr(34)&"javascript:document.form1.uname.value='"&rst("����")&"';"&chr(34)&">"&rst("����")&"</td><td>"&rst("�ȼ�")&"</td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write "<tr align=center><td colspan=2><form action=upgov.asp method=post name=form1>����:<input type=text size=14 maxlength=14 name='uname'> <input type=submit name='submit1' value=' �� �� '> <input type=submit name='submit1' value=' �� �� '>  <input type='submit' name='submit1' value=' �� �� '> <input type=submit name='submit1' value=' Ƹ �� '></form></td></tr></table></body>"
%>
