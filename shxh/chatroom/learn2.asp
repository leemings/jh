<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("Ba_jxqy_username")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
if instr(mg,"'")<>0 then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
co=Session("Ba_jxqy_usercorp")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select �ȼ�,���ľ���,��Ч from ���� where ����='"&un&"' and ��ʽ='"&mg&"'",conn
if rst.EOF or rst.BOF then
	rst.Close
	rst.Open "select * from ��ʽ where ����='"&co&"' and ��ʽ='"&mg&"'",conn
	if rst.EOF or rst.BOF then
		msg="<font color=FF0000>����ϰ��</font>##����ϰ"&mg&"�������޸���ָ�㣬���ʧ�ܡ�"
		learn="����û�д�����ʽ�ɹ���ϰ��"
	else
		energy=rst("���ľ���")
		proviso=rst("��ϰ����")
		basemp=rst("��������")
		baseat=rst("��������")
		especial=rst("��Ч")
		atdeclaration=rst("����˵��")
		rst.Close
		rst.Open "select * from �û� where ����='"&un&"' and ����>="&energy&" and "&proviso,conn
		if rst.EOF or rst.BOF then
			provisotxt=replace(proviso,"and","����")
			provisotxt=replace(provisotxt,"or","����")
			provisotxt=replace(provisotxt,">=","������")
			provisotxt=replace(provisotxt,"<=","������")
			provisotxt=replace(provisotxt,">","����")
			provisotxt=replace(provisotxt,"<","С��")
			provisotxt=replace(provisotxt,"=","Ϊ")
			msg="<font color=FF0000>����ϰ��</font>##��ϰ"&mg&"ʱ�����и���ָ�㣬�����Լ��������ޡ����ʧ�ܡ�"
			learn="��ϰʱʧ�ܣ��辫��"&energy&"���������㣺"&provisotxt
		else
			conn.Execute "insert into ����(����,��ʽ,�ȼ�,���ľ���,��������,��������,��Ч,����˵��) values('"&un&"','"&mg&"',1,"&energy&","&basemp&","&baseat&",'"&especial&"','"&atdeclaration&"')"
			conn.execute "update �û� set ����=����-"&energy&" where ����='"&un&"'"
			msg="<font color=FF0000>����ϰ��</font>##��ϰ"&mg&"��1���ɹ������þ���<bgsound src='../mid/xiulian.wav' loop=1>"&energy
			learn="��ϲ����ϰ"&mg&"���ţ��Ժ����������������"
		end if	
	end if
elseif rst("�ȼ�")<100 then
	energy=rst("���ľ���")
	agrade=rst("�ȼ�")
	rst.Close
	rst.Open "select * from �û� where ����='"&un&"' and ����>="&energy*agrade,conn
	if rst.EOF or rst.BOF then
		msg="<font color=FF0000>����ϰ��</font>##��ϰ"&mg&"��"&agrade&"��ʱ��������ʧ�ܣ��޷�ר���붨�����ʧ��"
		learn="��ϰʱʧ�ܣ��辫��"&energy
	else
		conn.Execute "update �û� set ����=����-"&energy*agrade&" where ����='"&un&"'"
		conn.Execute "update ���� set �ȼ�=�ȼ�+1,��������=��������*1.1 where ����='"&un&"' and ��ʽ='"&mg&"'"
		msg="<font color=FF0000>����ϰ��</font>##��ϰ"&mg&"��"&agrade&"���ɹ������þ���<bgsound src='../mid/xiulian.wav' loop=1>"&energy*agrade
		learn="��ϲ����ϰ"&mg&"��"&agrade+1&"���Ժ����������������"
	end if
else	msg="<font color=FF0000>����ϰ��</font>##���Ѿ���ϰ"&mg&"��������޷�����������<bgsound src='../mid/xiulian.wav' loop=1>"
	learn="����ȷ������"&mg&"�Ǹ����еĸ����ˣ�����ϰҲ������ʲôЧ���ˣ�"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=replace(msg,"\","\\")
msg=replace(msg,"/","\/")
msg=replace(msg,chr(34),"\"&chr(34))
msg=replace(msg,chr(39),"\"&chr(39))
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
newtalkarr(593)=1
newtalkarr(594)=un
newtalkarr(595)=un
newtalkarr(596)=""
newtalkarr(597)="#660099"
newtalkarr(598)="#660099"
newtalkarr(599)=msg&"<font class=timsty>��"&time()&"��<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<html>
<head>
<title>��ϰ�书</title>
<link rel=stylesheet href='css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<div align=center>
<font color=0000ff size=5>��ϰ�书</font>
<hr>
<input type=button value='����' onclick="javascript:location.href='learn.asp'">
</div>
<%=learn%>
</body>
</html>