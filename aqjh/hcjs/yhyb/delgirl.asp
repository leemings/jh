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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ID,���� FROM �û� WHERE ����='" & aqjh_name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�������Ҳ���������ϣ�');window.close();</script>"
	response.end
end if
aqjh_id=rs("id")
yinliang=rs("����")
rs.close
%><!--#include file="jiu.asp"--><%
sql="select * from ���� where ����ID="& aqjh_id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('��Ū���˰ɣ���Ժû������������ѽ��');window.close();</script>"
	response.end
end if
if yinliang<100000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('���������û���㹻��Ǯ�����������Ŷ�û��');window.close();</script>"
	response.end
end if
conn.execute "update �û� set ����=����-100000000 where ����='"& aqjh_name &"'"
connt.execute("delete * from ���� where ����ID="&aqjh_id)
rs.close
set rs=nothing
conn.close
set conn=nothing
connt.close
set connt=nothing
Response.Write "<script Language=Javascript>alert('��������������������������ˣ���ӭ����ûǮ��ʱ�����������⹤����');window.close();</script>"
response.end
%>
