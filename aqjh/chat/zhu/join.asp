<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="../../showchat.asp" -->
<%
aqjh_name=Session("aqjh_name")
username1=request("name1")
username2=request("name2")
if aqjh_name<>username1 then
if aqjh_name=username2 then
response.Write "<script language=JavaScript>alert('��ʾ�����Լ�������˼�����飬�Լ��Ѿ����鳤�ˣ�');window.close();</script>"
response.End()
else
response.Write "<script language=JavaScript>alert('��ʾ��������������˼���Ҫ����������');window.close();</script>"
response.End()
end if
end if
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ:�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
act=1
nowinroom=session("nowinroom")
addwordcolor="660099"
saycolor="008888"
towho="���"
addsays=aqjh_name

ffsj=application("aqjh_jointime")
nowsj=DateDiff("s",ffsj,now())
if nowsj>=30 then
        Application.Lock
	Application("aqjh_jointime")=""
        Application.UnLock
	Response.Write "<script language=JavaScript>{alert('�����齭��������ʾ�������������Чʱ��30�룬�������ϡ�');window.close();}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs1.open "select * from �û� where ����='"&username2&"'",conn,1,3
zid=rs1("id")
zname=rs1("����")
rs1.close
rs.open "select * from �û� where ����='"&username1&"'",conn,1,3
zdzt=rs("���")
zdm=rs("����")
if zdzt=false then
response.Write "<script language=JavaScript>alert('��ʾ���������״̬��û�򿪣���ô�ܼ��뵽���أ�');window.close();</script>"
response.End()
rs.close
conn.close
set rs=nothing
set conn=nothing
end if
if zdm<>"" and zdm<>"��" then
response.Write "<script language=JavaScript>alert('��ʾ�����Ѿ��������ˣ����ܼ��������飡');window.close();</script>"
response.End()
rs.close
conn.close
set rs=nothing
set conn=nothing
end if
rs("����")=zname
rs.update
rs.close
rs.open "select * from ��� where �鳤="&zid&" and ����='"&zname&"'",conn,1,3
if rs.eof and rs.bof then
response.Write "<script language=JavaScript>alert('��ʾ�������ڵ���!');window.close();</script>"
response.End()
rs.close
conn.close
set rs=nothing
set conn=nothing
end if 
zttname=rs("��Ա")
rs("��Ա")=zttname&","&username1
rs("��Ա��")=rs("��Ա��")+1
rs("ʱ��")=now()
rs.update
rs.close
conn.close
set rs=nothing
set conn=nothing
diaox="<font color=#ff0000>�������Ϣ��</font> <font color=#ff00ff>" & username1 & "</font>Ӧ�鳤<font color=#ff00ff>"&username2&"</font>������,��������["&zname&"]!"
call showchat(diaox)
Response.Write "<script language=JavaScript>{alert('������ɹ���');location.href = window.close();}</script>"
response.End()
%>
