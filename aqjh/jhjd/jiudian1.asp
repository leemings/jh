<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=request("id")
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
sql="insert into ����б�(�����,ӵ����,����,����,����,���,����,�ۼ�,ʱ��) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now())"
rs=connt.execute(sql)
connt.close
set connt=nothing
says="<font color=red>������Ϣ��</font><font color=green>"&my&"�ڽ�����Ƶ��"&wu&"������<font color=red>��"&lx&"��</font>��ᣬ��Ҷ���ȥѽ�����˾�û�ĳ��ˡ�����</font>"			'��������
call showchat(says)
Response.Redirect "jd.asp"
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
%>