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
on error resume next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn
if not(rst.EOF or rst.BOF) then Response.Redirect "../error.asp?id=053"
rst.Close
set rst=nothing
conn.BeginTrans
conn.Execute "Insert Into ��Ʊ(��Ʊ����,��ͨ������,ʣ��ɷ�,���м�,��ʷ�߼�,��ʷ�ͼ�,����ɽ���,�ּ�) values('"&stock&"',"&num&","&num&","&price&","&price&","&price&","&price&","&price&")"
if conn.Errors.Count>0 then
	conn.RollbackTrans
	Response.Redirect "../error.asp?id=104&errormsg="&conn.Errors(0).Description
else
	conn.CommitTrans
end if
conn.Close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><script language=javascript>setTimeout('history.go(-1);',3000);</script><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center><font color=0000ff size=4>�����¹�</font><br>�����¹����.�����Ӻ��Զ�����<br><a href='javascript:history.go(-1);'>����</a></p>"
%>