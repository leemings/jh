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

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,mvalue,w5,�ȼ�,��Ա�ȼ� FROM �û� WHERE ����='" & aqjh_name & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ڻ��������ǽ������˰ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("�ȼ�")<>55 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('Ϊ��ֹ���ף�ֻ�еȼ���55����ʱ��ſ�����ȡ��Ʒ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("mvalue")<=8000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�»�����8000���ϵĲſ�����ȡ��Ʒ��');location.href = 'javascript:history.go(-1)';</script>"
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
		jb=200
		sql="update �û� set ���=���+"&jb&" where id="&myid
		lw0="���"&jb&"��"

		jk=2
		sql1="update �û� set ��Ա��=��Ա��+"&jk&" where id="&myid
		lw1="��Ա��"&jk&"��"

		kp=2
		temp=add(wpdate,"����",kp)
		sql2="update �û� set w5='"&temp&"' where id="&myid
		lw="����"&kp&"��"
set rs=conn.execute(sql)
set rs=conn.execute(sql1)
set rs=conn.execute(sql2)
conn1.execute "insert into  ��������(yhid,����,����) values ('"&myid&"','"&aqjh_name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
conn1.close
set conn1=nothing
says="<font color=red><b>��������Ϣ��</b></font>"&aqjh_name&"��վ������õ�"&lw0&""&lw1&""&lw&""

call showchat(says)



%>
<html>
<head>
<title>���﷢��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body text="#000000" background="../chat/bg.gif">
<div align="center">
<table width="106" border="0" cellspacing="0" cellpadding="0" bordercolor="#0A246A" bgcolor="#808000">
  <tr>
    <td width="26">��ȡ�ɹ�</td>
    <td width="80" align="center"><b><font color="#FFFF00">��ϲ���<%=application("aqjh_user")%>����õ���</font><br><font color="#FF0000"><%=lw0%><%=lw1%><%=lw%></font></b></td>
  </tr>
</table>
</div>
</body>
</html>
</body>
</html>
