<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->

<%Response.charset="gb2312"%><%
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ���,����ʱ��,���� FROM �û� WHERE ����='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('�������["& ss &"]��,�ٲ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("����")<>"" and rs("����")<>"��" then 
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���������ڣ����뿪����ٽ��д˲�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if

if rs("���")=true then
	conn.Execute "update �û� set ���=false,����ʱ��=now() where ����='"&aqjh_name&"'"	
		diaox="��ֹ����ӣ�ϣ�����ҵ��鳤��Ҫɧ�������Ҽ���Ա��۾����"

else
	diaox="�����״̬�����ڿ��Խ��鳤������"
	conn.Execute "update �û� set ���=true,����ʱ��=now() where ����='"&aqjh_name&"'"	
end if
	Response.Write "<script language=JavaScript>{alert('���״̬�л��ɹ���');location.href = 'javascript:history.go(-1)';}</script>"

        rs.close

conn.close
set rs=nothing
set conn=nothing
Application.Lock

diaox="<font color=#ff0000>�������Ϣ��</font> <font color=#ff00ff>" & aqjh_name & "</font>"&diaox
call showchat(diaox)

%>

</body>
</html>