<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
corpid=Request.Form("corpid")
corp=Request.Form("corp")
earnings=Request.Form("earnings")
introduce=Request.Form("introduce")
condition=Request.Form("condition")
chaton=cbool(Request.Form("chaton"))
opt=Request.Form("opt")
if corpid="" or corp="" or earnings="" or introduce="" or condition="" or opt="" then Response.Redirect "..\errpr.asp?id=102"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ����  where id="&corpid,conn,1,2
if opt="����"  then 
	rst.AddNew
end if
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="ɾ��" then
	rst.Delete
else
	rst("����")=corp
	rst("����ϵ��")=earnings
	rst("���")=introduce
	rst("��������")=condition
	rst("chaton")=chaton
end if	
rst.Update
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
Response.write "<head><link rel='stylesheet' href='../chatroom/style1.css'>.</head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"'><div height=100% align=center valign=middle>�������<br><a href="&chr(34)&"javascript:location.replace('modcorp1.asp');"&chr(34)&">����</a></div></body>"
%>
