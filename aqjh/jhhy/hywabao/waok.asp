<%@ LANGUAGE=VBScript%>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	ss=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"�������ڱ��ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
randomize timer
r=int(rnd*6)+3
money=0
tl=0
nl=0
zstemp=rs("w6")
select case r
case 1
js(0) ="��ͷ"
js(1) ="����ʯ"
js(2) ="�챦ʯ"
js(3) ="����ʯ"
js(4) ="�̱�ʯ"
js(5) ="ˮ��ʯ"
js(6) ="�̱�ʯ"
js(7) ="���"
js(8) ="����"
js(9)="ľ̿"
	randomize()
	myxy = Int(Rnd*10)
	zstemp=add(rs("w6"),js(myxy),1)
	mess="��ϲ"& aqjh_name &"�ڵ�һ[<font color=blue>"&js(myxy)&"</font>]һ�飬����Ǹ��ö���ѽ��"
	erase js
case 2
	mess="��ϲ"& aqjh_name &"�ڵ�һ������������������200������"
	money=2000000
case 3
	mess="��ϲ"& aqjh_name &"�ڵ�һ��E���񽣣������õ�����500����"
	money=5000000
case 4
	mess="��ϲ"& aqjh_name &"�ڵ�һ��ħ����ָ�������õ�����100����"
	money=1000000
case 5
	mess="��ù��"& aqjh_name &"����û�ҵ������Ҳ�С�Ĳȵ�����,��������50000����������20000"
	tl=50000
	nl=20000
case 6
	mess="ǿ����������,"& aqjh_name &"�����⵽����,�����½�50000�������½�50000"
	tl=50000
	nl=50000
case 7
	rs.close
	rs.open "select ���� from �û� where ����='" & Application("aqjh_baowuname") & "'",conn
	if rs.eof or rs.bof  then
	    mess=""& aqjh_name &"�������Ǻã���Ȼ�ڳ��˽�������"&Application("aqjh_baowuname")&"����һ���������"
		conn.execute  "update �û� set ����=false,��������=0,����='"& Application("aqjh_baowuname") &"' where ����='" & aqjh_name &"'"
	else
		mess="ǿ����������,"& aqjh_name &"��ǿ�������ˣ��õ�����100��!"
		money=1000000
	end if
end select
conn.execute "update �û� set ����ʱ��=now(),����=����-"&tl&",����=����-"&nl&",����=����+"&money&",w6='"&zstemp&"'  where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>��Ϣ</font>"& aqjh_name &"����ɽ�ڱ���"&mess			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
<title>�ڱ���OK</title></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
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
</div></td></tr></table></div></body></html>