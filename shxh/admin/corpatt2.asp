<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=Session("Ba_jxqy_usercorp")
if username="" then Response.redirect "../error.asp?id=016"
attackname=Server.HTMLEncode (Trim(Request.Form("attackname")))
especial=Server.HTMLEncode (Trim(Request.Form("especial")))
energy=Trim(Request.Form("energy"))
if instr(attackname,"'")<>0 or instr(attackname," ")<>0 or instr(attackname,"\")<>0 or instr(attackname,"/")<>0 or instr(attackname,chr(34))<>0 then Response.Redirect "../error.asp?id=056"
if instr(";��;���;�ж�;���;����;",";"&especial&";")=0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(energy) then Response.Redirect "../error.asp?id=024"
energy=clng(energy)
if energy<10 then energy=10
if especial="��" then 
	mp=energy\10
else
	mp=energy
end if		
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"' and ����='"&usercorp&"' and ���='����' and ����>="&energy*20,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=065"
rst.Close
on error resume next
conn.BeginTrans
conn.Execute "insert into ��ʽ(��ʽ,����,��ϰ����,���ľ���,��������,��������,��Ч,˵��,����˵��) values('"&attackname&"','"&usercorp&"','True',"&energy&","&mp&","&energy&",'"&especial&"','"&username&"�������ʺ���������ϰ��','##��%%ʹ����"&usercorp&"�Ķ����书"&attackname&"')"
conn.Execute "update �û� set ����=����-"&energy*20&" where ����='"&username&"'"
if conn.Errors.Count>0 then
	conn.RollbackTrans
	Response.Redirect "../error.asp?id=104&errormsg="&conn.Errors(0).Description
else
	conn.CommitTrans
end if		
conn.Close
set conn=nothing
%>
<script language=javascript>history.back();</script>