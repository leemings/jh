<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=trim(request.form("id"))
qingren=trim(request.form("name"))
my=aqjh_name
%><!--#include file="dadata.asp"-->
<%
sql="select * from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "�㲻�ǽ������ˣ����ܶ�������"
conn.close
response.end
else
sex=rs("�Ա�")
if sex="��" then
sex="Ů"
else
sex="��"	
end if
sql="select * from �û� where ����='" & qingren & "' and �Ա�='"&sex&"' and ����<>'" & my & "' "
set rs=conn.execute(sql)
if rs.eof or rs.bof then
conn.close
Response.Write "<script language=javascript>alert('������û���������"&sex&"������!');history.back();</script>"
response.end
else
qingren=rs("����")
sql="SELECT * FROM ��� where ID=" & id
Set Rs=connt.Execute(sql)
wu=rs("�����")
yin=rs("�ۼ�")
lx=rs("����")
nl=rs("����")
tl=rs("����")
jb=rs("���")
sl=rs("����")%>
<%
sql="select * from �û� where ����='" & my & "'"
rs=conn.execute(sql)
if yin<=rs("����") then
sql="update �û� set ����=����-" & yin & " where ����='" & my & "'"
rs=conn.execute(sql)%>
<%
sql="select * from ����б� where �����='" & wu & "' and ӵ����='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into ����б�(�����,ӵ����,����,����,����,���,����,�ۼ�,�μ���,ʱ��) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&qingren&"',now())"
rs=connt.execute(sql)
connt.close
set connt=nothing
if Instr(Application("aqjh_useronlinename"&session("nowinroom"))," "&qingren&" ")<>0 then
says="<font color=red>������Ϣ��</font><font color=green>"&my&"�ڽ�����Ƶ��"&wu&"��Ϊ"&qingren&"����<font color=red>��"&lx&"��</font>��ᣬ"&qingren&"��ȥѽ���Ǻ�..</font>"		'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
call showchat(says)
end if
Response.Write "<script language=javascript>alert('��ϲ����Ϊ"&qingren&"���ľ����Ѿ�׼�����ˣ�');window.close();</script>"
response.end
else
connt.close
Response.Write "<script language=javascript>alert('���ܶ����磬ԭ�����Ѷ���ͬ�����ľ�ϯ��');history.back();</script>"
response.end
end if
else
connt.close
Response.Write "<script language=javascript>alert('���ܶ����磬ԭ���������������');history.back();</script>"
response.end
end if
rs.close
set rs=nothing
end if
end if
%>