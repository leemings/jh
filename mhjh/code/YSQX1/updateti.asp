<%
Response.Expires=0
bgcolor=Application("yx8_mhjh_backgroundcolor")
bgimage=Application("yx8_mhjh_backgroundimage")
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
if not(usergrade=120 and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
schid=Request.Form("schid")
question=Request.Form("question")
answer=Request.Form("answer")
ti=Request.Form("ti")
money=Request.Form("money")
qiang=cbool(Request.Form("qiang"))
if not isnumeric(schid) then Response.Redirect "../error.asp?id=024"
'for each element in Request.Form
'	if Request.Form(element)="" then Response.Redirect "../error.asp?id=102"
'next	
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background=../chatroom/bg1.gif bgcolor='"&bgcolor&"'><p align=center>����</p><hr>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("yx8_mhjh_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM ���� where id="&schid
rst.Open sqlstr,conn,1,2
if Request.Form("submit")="����" then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
if Request.Form("submit")="ɾ��" then 
	rst.Delete
	msg=msg&"<tr><td colspan=2>�ü�¼��ɾ��</td></tr>"
else
        rst("����")=question
	rst("��")=answer
	rst("�ṩ��")=ti
	rst("����")=money
	rst("����")=qiang
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"<p align=center><input type='button' value='��ȷ ����' onclick='javascript:history.go(-2)'></p></body>"
Response.Write msg
%>
