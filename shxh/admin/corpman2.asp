<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
if username="" then Response.redirect "../error.asp?id=016"
st=Trim(Request.Form("st"))
opt=Trim(Request.Form("submit"))
degree=Server.HTMLEncode(Trim(Request.Form("degree")))
if instr(st,"'")<>0 or instr(degree,"'")<>0 or degree="����" then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject ("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"' and ���='����'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=046"
corp=rst("����")
rst.Close
rst.Open "select * from �û� where ����='"&st&"' and ����='"&corp&"' and ���<>'����'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=046"
rst.Close
set rst=nothing
if opt="���" then
	conn.Execute "update �û� set ���='"&degree&"' where ����='"&st&"' and ����='"&corp&"' and ���<>'����'"
	msg="<font color=ff0000>����⡿</font>##��%%΢Ц��˵�����Ժ��������"&corp&"��"&degree&"�ˣ���Ҫ�úø�ѽ����"
else
	conn.Execute "update �û� set ���='��',����='��' where ����='"&st&"' and ����='"&corp&"' and ���<>'����'"
	msg="<font color=ff0000>�������ǽ��</font>##��%%���һ�������Ժ���"&corp&"û���������ĵ����ˣ��������������֮����������ǽ"
end if
conn.Close
set con=nothing
talkarr=Application("Ba_jxqy_talkarr")
talkpoint=clng(Application("Ba_jxqy_talkpoint"))
dim newtalkarr(600)
j=1
for i=11 to 600 step 10
	newtalkarr(j)=talkarr(i)
	newtalkarr(j+1)=talkarr(i+1)
	newtalkarr(j+2)=talkarr(i+2)
	newtalkarr(j+3)=talkarr(i+3)
	newtalkarr(j+4)=talkarr(i+4)
	newtalkarr(j+5)=talkarr(i+5)
	newtalkarr(j+6)=talkarr(i+6)
	newtalkarr(j+7)=talkarr(i+7)
	newtalkarr(j+8)=talkarr(i+8)
	newtalkarr(j+9)=talkarr(i+9)
	j=j+10
next
newtalkarr(591)=talkpoint+1
newtalkarr(592)=2
newtalkarr(593)=0
newtalkarr(594)=username
newtalkarr(595)=st
newtalkarr(596)=""
newtalkarr(597)="660099"
newtalkarr(598)="660099"
newtalkarr(599)=msg&"<font class=timsty>��"&time()&"��<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
%>
<script language=javascript>history.back();</script>