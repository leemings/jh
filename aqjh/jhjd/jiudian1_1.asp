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
%><!--#include file="dadata.asp"-->
<%
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ����ܶ�������"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "SELECT * FROM ��� where ID=" & id,conn
wu=rs("�����")
yin=rs("�ۼ�")
lx=rs("����")
nl=rs("����")
tl=rs("����")
jb=rs("���")
sl=rs("����")%>
<%
rs.close
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if yin>=rs("����") then
	rs.close
	set rs=nothing
	connt.close
	set connt=nothing
	conn.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('���ܶ����磬ԭ���������������');history.back();</script>"
	response.end
end if
	conn.execute "update �û� set ����=����-" & yin & " where ����='" & aqjh_name & "'"
	rs.close
	rs.open "select * from ����б� where �����='" & wu & "' and ӵ����='" & aqjh_name & "'",conn
if rs.eof or rs.bof then
	connt.execute "insert into ����б�(�����,ӵ����,����,����,����,���,����,�ۼ�,ʱ��) values ('"&wu&"','"&aqjh_name&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now())"
	rs.close
	set rs=nothing
	connt.close
	set connt=nothing
says="<font color=red>������Ϣ��</font><font color=green>"&aqjh_name&"�ڽ�����Ƶ��"&wu&"������<font color=red>��"&lx&"��</font>��ᣬ��Ҷ���ȥѽ�����˾�û�ĳ��ˡ�����</font>"			'��������
call showchat(says)
	Response.Redirect "jd.asp"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('���ܶ����磬ԭ�����Ѷ���ͬ�����ľ�ϯ��');history.back();</script>"
	response.end
end if
%>
