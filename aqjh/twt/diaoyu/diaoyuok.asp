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
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"twt/diaoyu/diao.asp")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<8 then
	ss=8-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"����������ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����ʱ��=now()  where ����='"&aqjh_name&"'"
randomize timer
r=int(rnd*6)+1
select case r
case 1
	mess="��ϲ"& aqjh_name &"����һ�������㣬����Ϊ����ʹ�ã�ɱ������1200������800"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"������",1)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 2
	mess="��ϲ"& aqjh_name &"����һ�������㣬����������5������"
	conn.execute "update �û� set ����=����+50000 where ����='"&aqjh_name&"'" 
case 3
	mess="��ϲ"& aqjh_name &"����һֻ����㣬���¿�����������300������50"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"�����",1)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 4
	mess="��ϲ"& aqjh_name &"����һ��С���㣬���¿�����������100������30"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"С����",1)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
case 5
	mess="������"& aqjh_name &"��û������ˤ��������������300����������100"
	conn.execute "update �û� set ����=����-300,����=����-100 where ����='"&aqjh_name&"'"
case 6
	mess= aqjh_name &"͵�����������������˷��֣�һ��Ź�������½�500�������½�200"
	conn.execute "update �û� set ����=����-500,����=����-200 where ����='"&aqjh_name&"'"
case 7
	mess="��ϲ"& aqjh_name &"�������Ǻõ�BiangBiang��ѽ�����������㡢����㡢С�����һ����"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"������",1)
	duyao=add(duyao,"�����",1)
	duyao=add(duyao,"С����",1)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>���ӱߴ�����</font>"& aqjh_name &"�ںӱߵ��㣺"&mess			'��������
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
<title>����OK!</title></head>
<body oncontextmenu=self.event.returnValue=false background="../../bg.gif">
<div align="center"> <br>
<br><table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr><td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr><td valign="top">
<div align="center">
<p><%=mess%></p>
</div></table>
<div align="center"><br>
<input type=button value=�رմ��� onClick='window.close()' name="button">
</div></td></tr></table></div></body></html>