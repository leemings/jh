<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if aqjh_jhdj<3 then 
	Response.Write "<script Language=Javascript>alert('����С��������԰�������������Ĺ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ID,�Ա�,����,��ż FROM �û� WHERE ����='" & aqjh_name & "'",conn
aqjh_id=rs("id")
sex=rs("�Ա�")
meimao=rs("����")
peiou=rs("��ż")
rs.close

if sex<>"Ů" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('ι����û�и������԰����ֻ�й���ģ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if meimao<1000 then 
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�������ֵ����1000����԰���գ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if peiou<>"��" and peiou<>"" and peiou<>"δ��" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��԰�����ѻ�Ů�ˣ��鷳��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%><!--#include file="jiu.asp"--><% 
sql="select * from ��Ů where ��ŮID="&aqjh_id
set rs=connt.execute(sql)
if rs.bof or rs.eof then
	sql="insert into ��Ů(��ŮID,����,��ò��) values (" & aqjh_id & ",'" & aqjh_name & "'," & meimao & ")"
	connt.execute sql
	connt.close
	set connt=nothing
	conn.execute "update �û� set ����=����+2000000 where ����='" & aqjh_name & "'"
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "��ϲ������ʽ��Ϊ��԰�Ĺ���ȵ���������1200000��"
	response.write "<br><a href=index.asp>����</a>"
	response.end
end if
response.write "���Ѿ��Ǳ���԰�Ĺ����ˣ���ô�����Ǽ�ѽ��"
response.write "<br><a href=index.asp>����</a>"
rs.close
set rs=nothing
conn.close
set conn=nothing
connt.close
set connt=nothing
%>