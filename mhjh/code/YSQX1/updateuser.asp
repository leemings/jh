<%
Response.Buffer =true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
userid=Request.Form("id")
if not isnumeric(userid) then Response.Redirect "../error.asp?id=024"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center>�û�����</p><hr><table width=80% border=3 align=center><tr bgcolor=FFFF00><td width='35%'>����</td><td>��ֵ</td></tr>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("yx8_mhjh_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM �û� where id="&userid
rst.Open sqlstr,conn,1,2
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
if Request.Form("submit")=" ɾ �� " then
mate=rst("��ż")
username=rst("����")
if mate<>"��" then conn.Execute "update �û� set ��ż='��' where ����='"&mate&"'"
conn.Execute "delete * from ��Ʒ where ������='"&username&"'"%>
<!--#include file="data.asp"-->
<%
'conn.Execute "delete * from ý�� where ������='"&username&"'"
sqlq="delete * from ý�� where ������='"&username&"'"
 Set Rs=connt.Execute(sqlq)
conn.Execute "delete * from ���� where ����='"&username&"'"
rst.Delete
msg=msg&"<tr><td colspan=2>�ü�¼��ɾ��</td></tr>"
else
for i=1 to rst.Fields.Count-1
msg=msg&"<tr><td>"&rst.Fields(i).Name&"("&rst.Fields(i).Type&")</td><td>"&Request.Form(i+1)&"</td></tr>"
select case rst.Fields(i).type
case 202,130
rst.Fields(i).Value=cstr(Request.Form(i+1))
case 3
rst.Fields(i).Value=clng(Request.Form(i+1))
case 7
rst.Fields(i).Value=cdate(Request.Form(i+1))
case 11
rst.Fields(i).Value=cbool(Request.Form(i+1))
case else
rst.Fields(i).Value=Request.Form(i+1)
end select
next
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table><p align=center><input type='button' value='��ȷ ����' onclick='javascript:history.go(-2)'></p></body>"
Response.Write msg
%>
