<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
uname=Request("uname")
if uname="" then Response.Redirect "../error.asp?id=024"
ucur=request("ucur")
if ucur="" then ucur="ҩƷ"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��Ʒ where ������='"&uname&"' and ����='"&ucur&"' and ����>0 order by ���� desc",conn,1,3
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><p align=center><h4>"&ucur&"������:<font color=FF0000>"&uname&"</font><h4></p><hr><table width=50% align=center border=3><tr bgcolor=FFFF00 align=center><td>����</td><td>����</td><td>����</td><td>����</td></tr>"
do while not(rst.EOF or rst.BOF)
	Response.Write "<tr><td>"&rst("����")&"</td><td>"&ucur&"</td><td align=right>"&rst("����")&"</td><td><a href='chgcur.asp?id="&rst("ID")&"&uname="&uname&"&ucur="&ucur&"'>����</a></td></tr>"
	rst.MoveNext 
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write "</table><p align=center><a href="&chr(34)&"javascript:location.replace('showuser.asp?search="&uname&"');"&chr(34)&">����</a></p></body>"
%>