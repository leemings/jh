<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<!--#include file="../showchat.asp"-->

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');</script>"
	Response.End
end if
	rq=Day(date())

if rq<>11 then

	Response.Write "<script Language=Javascript>alert('��ʦ�ڣ�9.10����û������������ȡ����,������"&rq&"�ţ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,mvalue,w5,�ȼ�,��Ա�ȼ�,ת�� FROM �û� WHERE ����='" & aqjh_name & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ڻ��������ǽ������˰ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
If rs("�ȼ�")<100 and rs("ת��")<2 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('ת��2�Σ����ҵȼ��ﵽ100��������ȡ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("mvalue")<=60000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�»�����60000���ϵĲſ�����ȡ��Ʒ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
myid=rs("id")
wpdate=rs("w5")
hydj=rs("��Ա�ȼ�")
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("sdlw.mdb")
rs1.open "select * from �������� where yhid="&myid,conn1
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���Ѿ���������ˣ����ﲻ�࣬���Ǹ�������Ҳ����ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if




rs1.close
lw=""
randomize timer
r=int(rnd*13)+1
select case r
	case 1 '���
		n=int(rnd*10)+2
		sql="update �û� set ���=���+"&n&" where id="&myid
		lw="���"&n&"��"
	case 2 '��
		n=int(rnd*5)+1
		sql="update �û� set ��Ա��=��Ա��+"&n&" where id="&myid
		lw="��Ա��"&n&"��"
	case 3 '����
		n=int(rnd*2)+1
		temp=add(wpdate,"����",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="����"&n&"��"
	case 4 '��֮��
		n=int(rnd*2)+1
		temp=add(wpdate,"��֮��",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="��֮��"&n&"��"
	case 5 '�ֻ���
		n=int(rnd*2)+1
		temp=add(wpdate,"�ֻ���",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="�ֻ���"&n&"��"
	case 6 '���׿�
		n=int(rnd*1)+1
		temp=add(wpdate,"���׿�",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="���׿�"&n&"��"
	case 7 '������
		n=int(rnd*2)+1
		temp=add(wpdate,"������",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="������"&n&"��"
	case 8 '��ѥ��
		n=int(rnd*1)+1
		temp=add(wpdate,"��ѥ��",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="��ѥ��"&n&"��"
	case 9 '�氮��
		n=int(rnd*2)+1
		temp=add(wpdate,"�氮��",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="�氮��"&n&"��"
	case 10 '�ӵ㿨
		n=int(rnd*2)+1
		temp=add(wpdate,"�ӵ㿨",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="�ӵ㿨"&n&"��"
	case 11 'ս��
		n=int(rnd*2)+1
		temp=add(wpdate,"ս��",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="ս��"&n&"��"
	case 12 '������
		n=int(rnd*2)+1
		temp=add(wpdate,"������",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="������"&n&"��"
	case 13 '�񽣿�
		n=int(rnd*2)+1
		temp=add(wpdate,"�񽣿�",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="�񽣿�"&n&"��"	
end select
set rs=conn.execute(sql)
conn1.execute "insert into  ��������(yhid,����,����) values ('"&myid&"','"&aqjh_name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
conn1.close
set conn1=nothing
says="<font color=red><b>���������</b></font>"&aqjh_name&"��վ������õ���������"&lw&""

call showchat(says)

%>
<html>
<head>
<title>���﷢��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../chat/bg.gif">
<div align="center">
<table width="141" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="141" align="center"><b><font color="#FFFF00">��ϲ���<%=application("aqjh_user")%>����õ���</font><br><font color="#FF0000"><%=lw%></font></b></td>
  </tr>
</table>
</div>
</body>
</html>
</body>
</html>
