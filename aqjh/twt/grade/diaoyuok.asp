<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
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
if sj<2 then
	ss=2-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"������Ѱ���ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����ʱ��=now()  where ����='"&aqjh_name&"'"
randomize timer
r=int(rnd*18)+1
select case r
case 1
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ������ûѰ����ȥ���㣬�Ǻǣ��ջ��治���õ�����������"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"������",5)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 2
	mess= aqjh_name &"�ڻƽ�Ѱ������Ȼ�������������ܵ�����������������½�5000�������½�2800"
	conn.execute "update �û� set ����=����-5000,����=����-2800 where ����='"&aqjh_name&"'"
case 3
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����õ�ҩũ����������ûѰ����ȴ�������Ƶ���99��"
	rs.open "SELECT w1 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w1"),"���Ƶ���",99)
	conn.execute "update �û� set  w1='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 4
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����õ�ҩũ����������ûѰ����ȴ����Ů����99��"
	rs.open "SELECT w1 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w1"),"Ů����",99)
	conn.execute "update �û� set  w1='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 5
	mess= aqjh_name &"�ڻƽ�Ѱ�����ҵ�����ķ��������Ľ�⣬�õ����1��"
	conn.execute "update �û� set ���=���+1 where ����='"&aqjh_name&"'"
case 6
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����õ�ҩũ����������ûѰ����ȴ������ҩ99��"
	rs.open "SELECT w1 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w1"),"��ҩ",99)
	conn.execute "update �û� set  w1='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 7
	mess= aqjh_name &"�ڻƽ�Ѱ������С�ĵ��ºӣ������½�4900�������½�2400"
	conn.execute "update �û� set ����=����-4900,����=����-2400 where ����='"&aqjh_name&"'"
case 8
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����ں���ʰ��2�����"
	conn.execute "update �û� set ���=���+2 where ����='"&aqjh_name&"'"
case 9
	mess= aqjh_name &"�ڻƽ�Ѱ�������ϻ�Ϯ���������½�3800�������½�2400"
	conn.execute "update �û� set ����=����-3800,����=����-2400 where ����='"&aqjh_name&"'"
case 10
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ��������Ѱ����������������׽�㣬׽��һ�������"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"�����",1)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 11
	mess= aqjh_name &"�ڻƽ�Ѱ��,û�뵽��ǿ�����٣�һ�ٱ��������½�3700�������½�3200"
	conn.execute "update �û� set ����=����-3700,����=����-3200 where ����='"&aqjh_name&"'"
case 12
	mess= aqjh_name &"�ڻƽ�Ѱ����˭֪�������ˣ�һ��ȭ����ߣ������½�2800�������½�1000"
	conn.execute "update �û� set ����=����-2800,����=����-1000 where ����='"&aqjh_name&"'"
case 13
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����õ���ũ����������ûѰ����ȴ���������õ��99��<img src='../hcjs/jhjs/images/flower390.gif'>"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"�����õ��",99)
	conn.execute "update �û� set  w7='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 14
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����õ���ũ����������ûѰ����ȴ������������99��<img src='../hcjs/jhjs/images/f24.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"��������",99)
	conn.execute "update �û� set  w7='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 15
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����õ���ũ����������ûѰ����ȴ����ҡͷ��õ��99��<img src='../hcjs/jhjs/images/flower370.gif'>"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"ҡͷ��õ��",99)
	conn.execute "update �û� set  w7='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 16
	mess="��λ˧����Ů!!"& aqjh_name &"�ڻƽ�Ѱ�����õ���ũ����������ûѰ����ȴ������������99��<img src='../hcjs/jhjs/images/flowerr12.GIF'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"��������",99)
	conn.execute "update �û� set  w7='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 17
	mess= aqjh_name &"�ڻƽ�Ѱ����˭֪�������ˣ�һ��ȭ����ߣ������½�2800�������½�1000"
	conn.execute "update �û� set ����=����-2800,����=����-1000 where ����='"&aqjh_name&"'"
case 18
	mess= aqjh_name &"�ڻƽ�Ѱ������Ȼ�������������ܵ�����������������½�5000�������½�2800"
	conn.execute "update �û� set ����=����-5000,����=����-2800 where ����='"&aqjh_name&"'"
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>���ƽ�Ѱ����</font>"& aqjh_name &"Ѱ����"&mess			'��������
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
<link href="../../css.css" rel=stylesheet type="text/css">
<title>Ѱ��OK!</title></head>
<body oncontextmenu=self.event.returnValue=false background="../../bg.gif">
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
</table></div></body></html>