<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.redirect "../error.asp?id=016"
corpcond=server.HTMLEncode(Trim(Request.Form("corpcond")))
if instr(corpcond,"'")<>0 or instr(corpcond,chr(34))<>0 then Response.Redirect "../error.asp?id=056"
tj=Request.Form("corptj")
corpname=server.HTMLEncode(Trim(Request.Form("corpname")))
if corpname="" or corpname="��" or instr(corpname,"'")<>0 or instr(corpname,chr(34))<>0 or instr(corpname," ")<>0 or instr(corpname,"�")<>0 or instr(corpname,"&#63733;") then Response.Redirect "../error.asp?id=056"
corpsilver=server.HTMLEncode(Trim(Request.Form("corpsilver")))
if not isnumeric(corpsilver) then Response.Redirect "../error.asp?id=024"
silver=clng(corpsilver)
if silver<0 then 
	silver=0
elseif silver>1000 then
	silver=1000
end if
if tj="" then Response.Redirect "../error.asp?id=102"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"' and ����<>'�ٸ�' and ���<>'����' and ����>100000 and ����>100000 and ����>100000 and ����>100000 and ����>"&silver*100000&" and ����>10000",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=063"
rst.Close
rst.Open "select * from ���� where ����='"&corpname&"'",conn,1,2
if rst.EOF or rst.BOF then
	rst.AddNew
	rst("����")=corpname
	rst("����ϵ��")=silver
	rst("���")=corpcond
	rst("��������")=tj
	rst("chaton")=false
	rst.Update
	conn.Execute "update �û� set ����=����-10000,����=����-"&silver*100000&",���='����',����='"&corpname&"' where ����='"&username&"'"
	Session("Ba_jxqy_usercorp")=corpname
else 
	Response.Redirect "../error.asp?id=064"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing 
%>
<head><link rel='stylesheet' href='../chatroom/style1.css'></head><body oncontextmenu=self.event.returnValue=false background='<%=bgimage%>' bgcolor='<%=bgcolor%>' topmargin=100><p align=center><font color=ff0000 size=5>�����Ż�</font></p><hr>
��ϲ��ϲ���������ڵķܶ�,�����ڽ������Լ�������<font color=0000ff size=6><%=corpname%></font> ��������10000,��������<%=silver*100000%><br><p align=center><a href=javascript:location.replace('addcorp.asp');>����</a></p>
