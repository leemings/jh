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
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���,��Ա��,w6,����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 then
	ss=3-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"��������������');window.close();</script>"
	Response.End
end if
randomize timer
r=int(rnd*24)+1
jinbi=0
jinka=0
zstemp=rs("w6")
select case r
case 1
	dim js(10)
	js(0) ="��ͷ"
	js(1) ="����ʯ"
	js(2) ="�챦ʯ"
	js(3) ="��ʬѪ"
	js(4) ="��ʬ��"
	js(5) ="ˮ��ʯ"
	js(6) ="�̱�ʯ"
	js(7) ="���"
	js(8) ="����"
	js(9)="ľ̿"
	randomize()
	myxy = Int(Rnd*10)
	zstemp=add(rs("w6"),js(myxy),1)
	mess="��ϲ"& aqjh_name &"��һ[<font color=blue>"&js(myxy)&"</font>]һ�飬����Ϊ��ҩϵͳ�ر��ĺö��� ��"
	sql="update �û� set ����ʱ��=now(),w6='"&zstemp&"' where ����='"&aqjh_name&"'"
	erase js
case 2
	mess=""& aqjh_name &"������ã��񵽱��˲�С�ĵ��Ľ��<img src='img/jinbi.gif'>1��"
	sql="update �û� set ����ʱ��=now(),���=���+1 where ����='"&aqjh_name&"'"
case 3
	mess=""& aqjh_name &"������ã��񵽱��˲�С�ĵ��Ľ��<img src='img/jinbi.gif'>2��"
	sql="update �û� set ����ʱ��=now(),���=���+2 where ����='"&aqjh_name&"'"
case 4
	mess=""& aqjh_name &"������ã��񵽱��˲�С�ĵ��Ľ��<img src='img/jinbi.gif'>3��"
	sql="update �û� set ����ʱ��=now(),���=���+3 where ����='"&aqjh_name&"'"
case 5
	mess=""& aqjh_name &"������ã������˽�������˱��ˣ������������<img src='img/jinbi.gif'>2��"
	sql="update �û� set ����ʱ��=now(),���=���+2 where ����='"&aqjh_name&"'"
case 6
	mess=""& aqjh_name &"�������,����һ��Ů��������������,���������Ź��,���¼���1000"
	sql="update �û� set ����ʱ��=now(),����=����-1000 where ����='"&aqjh_name&"'"
case 7
	mess=""& aqjh_name &"�������,��������˺ͽ������Ӽ������ǣ�����Ź��һ��,��������1500"
	sql="update �û� set ����ʱ��=now(),����=����-1500 where ����='"&aqjh_name&"'"
case 8
	mess=""& aqjh_name &"������ã������˽����׸��ܰ�Ƥ��ƽ�����ܰ�Ƥ��С����һë�����ˣ����첻֪����ʲôϲ���ˣ������ر�ã����������<img src='img/jinbi.gif'>4��"
	sql="update �û� set ����ʱ��=now(),���=���+4 where ����='"&aqjh_name&"'"
case 9
	mess=""& aqjh_name &"�������,�����������׷����С�ӣ�����Ź��һ��,��ȥ����һ���ֵܱ���,���2�ܾ���,��������1000��"
	sql="update �û� set ����ʱ��=now(),����=����-1000 where ����='"&aqjh_name&"'"
case 10
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,��������ܻ�ȥ����!~~"
        sql="update �û� set ����ʱ��=now() where ����='"&aqjh_name&"'"
case 11
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,�����½�100��"
	sql="update �û� set ����ʱ��=now(),����=����-100 where ����='"&aqjh_name&"'"
	case 12
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,�书��ʧ1000��"
	sql="update �û� set ����ʱ��=now(),�书=�书-1000 where ����='"&aqjh_name&"'"
	case 13
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,���㱻������10��"
	sql="update �û� set ����ʱ��=now(),�ݶ�����=�ݶ�����-10 where ����='"&aqjh_name&"'"
	case 14
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,������ʧ200��"
	sql="update �û� set ����ʱ��=now(),����=����-200 where ����='"&aqjh_name&"'"
	case 15
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,�õ�ҽ�Ʒ���2000Ԫ"
        sql="update �û� set ����ʱ��=now(),���=���+2000 where ����='"&aqjh_name&"'"
	case 16
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,����վ��·��,�Ȼ�һ����,ʲôҲû����ʧ"
	sql="update �û� set ����ʱ��=now() where ����='"&aqjh_name&"'"
	case 17
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,��һ��ü����,�������300��"
	sql="update �û� set ����ʱ��=now(),����=����-300  where ����='"&aqjh_name&"'"
	case 18
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,��һ��ü����,��������,������ƴ��,���������ʧ3000��,"
	sql="update �û� set ����ʱ��=now(),����=����-3000 where ����='"&aqjh_name&"'"
	case 19
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,�Լ��ܻؼ�,����3��,���û��������˺�,���Ǽ�������,������ʧ300��"
	sql="update �û� set ����ʱ��=now(),����=����+2000 where ����='"&aqjh_name&"'"
	case 20
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,�ؼҺ󷢷�ͼǿ,�����书,���ڳ�Ϊһ������,�书����2000��"
	sql="update �û� set ����ʱ��=now(),�书=�书+2000 where ����='"&aqjh_name&"'"
	case 21
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ�����Ź��һ��,��ɲз�,��������10000"
	sql="update �û� set ����ʱ��=now(),����=����-1000 where ����='"&aqjh_name&"'"
        case 22
	mess=""& aqjh_name &"������ã������˽�������˱��ˣ�����������<img src='img/jk.gif'>1Ԫ"
	sql="update �û� set ����ʱ��=now(),��Ա��=��Ա��+1 where ����='"&aqjh_name&"'"
	case 23
	mess=""& aqjh_name &"�������,��������˺ͽ��������������ǣ���㱻��Ź��һ��,�������书��ǿ,�ѻ��˴���仯��ˮ.��������500��"
	sql="update �û� set ����ʱ��=now(),����=����+500 where ����='"&aqjh_name&"'"
        case 24
        rs.close
	rs.open "select ���� from �û� where ����='"&aqjh_name&"' and ����='" & Application("aqjh_baowuname") & "'",conn
	if rs.eof or rs.bof  then
	    mess=""& aqjh_name &"�������Ǻã���Ȼ���˽�������"&Application("aqjh_baowuname")&"����һ���������"
            sql="update �û� set ����=false,��������=0,����='"& Application("aqjh_baowuname") &"' where ����='"&aqjh_name &"'"
        else
            mess=""& aqjh_name &"��Ȼ��ǰһ������Ȼ����ǰ�治Զ���ֽ�������"&Application("aqjh_baowuname")&"����׼��ȥ�ã�����̫̰�ģ��Լ��ı����Ѿ���һ�������ߡ�"
            sql="update �û� set ��������=0,����='��' where ����='"&aqjh_name &"'"
	end if
end select
conn.execute ""&sql&""
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>��������Ϣ��</font>"&mess			'��������
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
<link href="jh.css" rel=stylesheet type="text/css">
<title>������OK</title></head>
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
</div>
</td>
</tr>
</table>
</div>
</body></html>