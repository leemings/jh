<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
uname=Request("uname")
if uname="" then Response.Redirect "../error.asp?id=024"
ucur=request("ucur")
if ucur="" then ucur="ҩƷ"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��Ʒ where ������='"&uname&"' and ����='"&ucur&"' and ����>0 order by ���� desc",conn,1,3
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif' text='FF0000'><p align=center><h4>"&ucur&"������:<font color=FF0000>"&uname&"</font><h4></p><hr><table width=50% align=center border=3><tr bgcolor=FFFF00 align=center><td>����</td><td>����</td><td>����</td><td>����</td></tr>"
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
