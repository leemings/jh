<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	ss=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"�������ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����ʱ��=now()  where ����='"&aqjh_name&"'"
randomize timer
r=int(rnd*21)+1
select case r
case 1
	mess= aqjh_name &"��ʻ�Ʒɻ��ɹ����õ���������300��"
	conn.execute "update �û� set ����=����+300 where ����='"&aqjh_name&"'"
case 2
	mess= aqjh_name &"��ʻ40������Ʒɻ�����Ȼ�����������ɻ�������,�����½�50�������½�28��"
	conn.execute "update �û� set ����=����-500000,����=����-280000 where ����='"&aqjh_name&"'"
case 3
	mess= aqjh_name &"��ʻ�Ʒɻ�ʧ�£�����̫ƽ����ȥ���ϵ���"
	conn.execute "update �û� set ״̬='��',��¼=now(),�¼�ԭ��='�ϵ�|"&fn1&"' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'�ϵ�','�ɻ�ʧ��','����')"
case 4
	mess= aqjh_name &"��ʻ40������Ʒɻ��ɹ����õ���������200��"
	conn.execute "update �û� set ����=����+200 where ����='"&aqjh_name&"'"
case 5
	mess= aqjh_name &"��ʻ40������Ʒɻ�������̫ƽ���µ��죬�õ������ֵ����250��"
	conn.execute "update �û� set ����=����+250 where ����='"&aqjh_name&"'"
case 6
	mess= aqjh_name &"��ʻ40������Ʒɻ�����С�ĵ��ºӣ������½�49�������½�24��"
	conn.execute "update �û� set ����=����-490000,����=����-240000 where ����='"&aqjh_name&"'"
case 7
	mess="��λ��λ!!"& aqjh_name &"�������Ʒɻ������ˣ���������1000��"
	conn.execute "update �û� set ����=����+10000000 where ����='"&aqjh_name&"'" 
case 8
	mess= aqjh_name &"��ʻ�Ʒɻ��뻷�����У����Ϸɻ�ʧ�£����úòҰ�"
	conn.execute "update �û� set ״̬='��',��¼=now(),�¼�ԭ��='�ϵ�|"&fn1&"' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'�ϵ�','�ɻ�ʧ��','����')"
case 9
	mess= aqjh_name &"��ʻ40������Ʒɻ�����С��ײ�з���ʧ���ˣ������½�38�������½�24��"
	conn.execute "update �û� set ����=����-380000,����=����-240000 where ����='"&aqjh_name&"'"
case 10
	mess= aqjh_name &"��ʻ40������Ʒɻ����Խ������գ����ϱ����յ������У������½�37�������½�32��"
	conn.execute "update �û� set ����=����-3700,����=����-3200 where ����='"&aqjh_name&"'"
case 11
	mess= aqjh_name &"��ʻ40������Ʒɻ���˭֪�ɻ�ײ��ɭ�֣������½�28�������½�10��"
	conn.execute "update �û� set ����=����-280000,����=����-100000 where ����='"&aqjh_name&"'"
case 12
	mess= aqjh_name &"��ʻ40������Ʒɻ�������̫ƽ���µ��죬�õ������ֵ����250��"
	conn.execute "update �û� set ����=����+250 where ����='"&aqjh_name&"'"
case 13
	mess= aqjh_name &"��ʻ40������Ʒɻ�������̫ƽ���µ��죬�õ������ֵ����150��"
	conn.execute "update �û� set ����=����+150 where ����='"&aqjh_name&"'"
case 14
	mess= aqjh_name &"��ʻ�Ʒɻ�ʧ�£����Ϲ��°��찧��"
	conn.execute "update �û� set ״̬='��',��¼=now(),�¼�ԭ��='�ϵ�|"&fn1&"' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'�ϵ�','�ɻ�ʧ��','����')"
case 15
	mess="��λ��λ!!"& aqjh_name &"�������Ʒɻ������ˣ���������800��"
	conn.execute "update �û� set ����=����+8000000 where ����='"&aqjh_name&"'"
case 16
	mess="��λ��λ!!"& aqjh_name &"�������Ʒɻ������ˣ���������1000��"
	conn.execute "update �û� set ����=����+10000000 where ����='"&aqjh_name&"'"  
case 17
	mess= aqjh_name &"��ʻ�Ʒɻ�ʧ�£�����������ֻ��һ������"
	conn.execute "update �û� set ״̬='��',��¼=now(),�¼�ԭ��='�ϵ�|"&fn1&"' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'�ϵ�','�ɻ�ʧ��','����')"
case 18
	mess= aqjh_name &"��ʻ40������Ʒɻ�,��ϧ�ɻ����������ˣ������Ҫ����1000��"
	conn.execute "update �û� set ����=����-10000000 where ����='"&aqjh_name&"'"
case 19
	mess= aqjh_name &"��ʻ40������Ʒɻ�,��ϧ�ɻ����������ˣ������Ҫ����1000��"
	conn.execute "update �û� set ����=����-10000000 where ����='"&aqjh_name&"'"
case 20
	mess= aqjh_name &"���Լ����Ʒɻ���������ݣ��õ����1�������ǿ���"
	conn.execute "update �û� set ���=���+1 where ����='"&aqjh_name&"'"
case 21
	mess= aqjh_name &"��ʻ�Ʒɻ�ʧ�£���λ��ϺĬ��һ���Ӱ�"
	conn.execute "update �û� set ״̬='��',��¼=now(),�¼�ԭ��='�ϵ�|"&fn1&"' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'�ϵ�','�ɻ�ʧ��','����')"
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>��������Ϣ��</font>"& aqjh_name &"���ɻ���"&mess			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../dg/setup.css" rel=stylesheet type="text/css">
<title>����OK!</title></head>

<body oncontextmenu=self.event.returnValue=false background="../jhimg/bk_hc3w.gif">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=�رմ��� onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>

