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
	weekdate=weekday(date())
if weekdate<>6 then
	Response.Write "<script Language=Javascript>alert('��������ֻ��������ȡ���������"&weekdate-1&"��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,mvalue,w7,�ȼ�,��Ա�ȼ�,ת�� FROM �û� WHERE ����='" & aqjh_name & "'",conn
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
if rs("mvalue")<=20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�»�����20000���ϵĲſ�����ȡ��Ʒ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
myid=rs("id")
wpdate=rs("w7")
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
r=int(rnd*5)+1
select case r
	case 1 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=1000*hydj
                p=7000*hydj
		temp=add(wpdate,"��������",n)
		sql="update �û� set ��=��+"&m&",ľ=ľ+"&m&",ˮ=ˮ+"&m&",��=��+"&m&",��=��+"&k&",����=����+"&l&",�Ṧ=�Ṧ+"&m&",�书=�书+"&p&",w7='"&temp&"' where id="&myid
		lw="��������"&n&"��,��"&m&",ľ"&k&",ˮ"&k&",��"&m&",��"&k&",����"&l&",�Ṧ"&l&",�书"&p&""
	case 2 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=1000*hydj
                p=80000*hydj
		temp=add(wpdate,"��õ��",n)
		sql="update �û� set ��=��+"&m&",ľ=ľ+"&m&",ˮ=ˮ+"&m&",��=��+"&m&",��=��+"&k&",����=����+"&l&",�Ṧ=�Ṧ+"&m&",�书=�书+"&p&",w7='"&temp&"' where id="&myid
		lw="��õ��"&n&"��,��"&m&",ľ"&k&",ˮ"&k&",��"&m&",��"&k&",����"&l&",�Ṧ"&l&",�书"&p&""	
	case 3 
		n=999*hydj
                m=400*hydj
                k=600*hydj
                l=800*hydj
                p=6000*hydj
		temp=add(wpdate,"ȵ�����",n)
		sql="update �û� set ��=��+"&m&",ľ=ľ+"&m&",ˮ=ˮ+"&m&",��=��+"&l&",��=��+"&k&",����=����+"&l&",�Ṧ=�Ṧ+"&m&",�书=�书+"&p&",w7='"&temp&"' where id="&myid
		lw="ȵ�����"&n&"��,��"&m&",ľ"&k&",ˮ"&k&",��"&l&",��"&k&",����"&l&",�Ṧ"&l&",�书"&p&""
       case 4 
		n=999*hydj
                m=700*hydj
                k=600*hydj
                l=500*hydj
                p=6000*hydj
		temp=add(wpdate,"õ������",n)
		sql="update �û� set ��=��+"&m&",ľ=ľ+"&m&",ˮ=ˮ+"&m&",��=��+"&l&",��=��+"&k&",����=����+"&l&",�Ṧ=�Ṧ+"&m&",�书=�书+"&p&",w7='"&temp&"' where id="&myid
		lw="õ������"&n&"��,��"&m&",ľ"&k&",ˮ"&k&",��"&l&",��"&k&",����"&l&",�Ṧ"&m&",�书"&p&""
	case 5 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=900*hydj
                p=7000*hydj
		temp=add(wpdate,"�����õ��",n)
		sql="update �û� set ��=��+"&m&",ľ=ľ+"&m&",ˮ=ˮ+"&m&",��=��+"&k&",��=��+"&l&",����=����+"&k&",�Ṧ=�Ṧ+"&m&",�书=�书+"&p&",w7='"&temp&"' where id="&myid
		lw="�����õ��"&n&"��,��"&m&",ľ"&k&",ˮ"&k&",��"&k&",��"&k&",����"&k&",�Ṧ"&l&",�书"&p&""	
	case 6 
		n=999*hydj
                m=600*hydj
                k=500*hydj
                l=700*hydj
                p=5000*hydj
		temp=add(wpdate,"��ɵ�õ��",n)
		sql="update �û� set ��=��+"&m&",ľ=ľ+"&m&",ˮ=ˮ+"&m&",��=��+"&l&",��=��+"&m&",����=����+"&k&",�Ṧ=�Ṧ+"&m&",�书=�书+"&p&",w7='"&temp&"' where id="&myid
		lw="��ɵ�õ��"&n&"��,��"&m&",ľ"&k&",ˮ"&k&",��"&l&",��"&m&",����"&k&",�Ṧ"&l&",�书"&p&""
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
