<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
stock=Trim(Request.Form("stock"))
price=Request.Form("price")
num=Request.Form("num")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "��Ʊ",conn
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' ><p align=center ><font color=0000ff size=4>��Ʊ�ֺ�</font></p><hr><form action=allocatenow.asp method=post id=form1 name=form1><table width=60% align=center border=3><tr><td>��Ʊ����</td><td><select name='stock' size='10'>"
do while not (rst.EOF or rst.BOF)
	msg=msg&"<option value='"&rst("��Ʊ����")&"'>"&rst("��Ʊ����")&"</option>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
msg=msg&"</select></td></tr><tr><td>ÿ����Ϣ</td><td><input type=text name='money' value='5' size=3 maxlength=3></td></tr><tr align=center><td colspan=2><input type=submit name=opt value=' �� �� '> <input type=button onclick=javascript:history.back(); value=' �� �� ' id=button1 name=button1></td></tr></table></form></body>"
Response.Write msg
%>