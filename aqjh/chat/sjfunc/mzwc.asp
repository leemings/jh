<%@ LANGUAGE=VBScript%>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
rs.open "select ���,����,����,����,����,֪��,��Ա��,w6,ħ��,��,ľ,ˮ,��,��,����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	ss=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"�������ύ����');window.close();</script>"
	Response.End
end if
if aqjh_jhdj<90 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ȼ���Ҫ����90���ſ��ύ����');window.close();</script>"
	Response.End
end if
input=request.form("input")
if rs("ľ")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��û��500ľ�����޷��������');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("��")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��û��500�������޷��������');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("ħ��")<5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('����ħ���Ҳ����ô��ħ����û���أ�~��Ҫ5�㣬��ȥ�����ɣ�');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("ˮ")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��û��500ˮ�����޷��������');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("��")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��û��500�������޷��������');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("��")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��û��500�������޷��������');location.href = 'rwbz.asp';}</script>"
	response.end
end if
if trim(rs("����"))<>"ħ������" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������������ħ���������������ܽ���ħ�������������');window.close();</script>"
	response.end
end if
if rs("����")<1000000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�뽻1000000�������ύ�ѣ�');window.close();}</script>"
Response.End
end if
if rs("����")<100000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('������Ҫ100000������ͷ��ʲô����Ҫ������..');window.close();}</script>"
Response.End
end if
if rs("����")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('��Ҫ1000�����ſ����������');window.close();}</script>"
Response.End
end if
if rs("֪��")<200 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('������������Ͻ�200֪�ʣ�');window.close();}</script>"
Response.End
end if
conn.execute "update �û� set ����=����-10000 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ����=����-1000 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ����=����-1000 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ֪��=֪��-200 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ľ=ľ-500 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ��=��-500 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ˮ=ˮ-500 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ��=��-500 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ��=��-500 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ħ��=ħ��-5 where ����='"& aqjh_name &"'"
randomize timer
r=int(rnd*10)+1
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
	zstemp=add(rs("w6"),js(myxy),5)
	mess="˥��!!!!!"& aqjh_name &"�������õ�[<font color=blue>"&js(myxy)&"</font>]5�飬����Ϊ��ҩϵͳ�ر��ĺö��� ��"
	sql="update �û� set ����ʱ��=now(),w6='"&zstemp&"' where ����='"&aqjh_name&"'"
	erase js
case 2
	mess=""& aqjh_name &"��������书���1200���������500"
	sql="update �û� set ����ʱ��=now(),�书=�书+1200,����=����+500 where ����='"&aqjh_name&"'"
case 3
	mess=""& aqjh_name &"��������书���1000"
	sql="update �û� set ����ʱ��=now(),�书=�书+1000 where ����='"&aqjh_name&"'"
case 4
	mess=""& aqjh_name &"��������书+1300"
	sql="update �û� set ����ʱ��=now(),�书=�书+1300 where ����='"&aqjh_name&"'"
case 5
	mess=""& aqjh_name &"��������书+2000"
	sql="update �û� set ����ʱ��=now(),�书=�书+2000 where ����='"&aqjh_name&"'"
case 6
	mess=""& aqjh_name &"������񹥻�+120"
	sql="update �û� set ����ʱ��=now(),����=����+120 where ����='"&aqjh_name&"'"
	case 7
	mess=""& aqjh_name &"������񹥻�+150"
	sql="update �û� set ����ʱ��=now(),����=����+150 where ����='"&aqjh_name&"'"
	case 8
	mess=""& aqjh_name &"������񹥻�+180"
	sql="update �û� set ����ʱ��=now(),����=����+180 where ����='"&aqjh_name&"'"
	case 9
	mess=""& aqjh_name &"����������ӽ����������16660���п���������Ŷ��"
	sql="update �û� set ����ʱ��=now(),allvalue=allvalue+16660 where ����='"&aqjh_name&"'"
        case 10
        rs.close
	rs.open "select ���� from �û� where ����='"&aqjh_name&"' and ����='" & Application("aqjh_baowuname") & "'",conn
	if rs.eof or rs.bof  then
	    mess=""& aqjh_name &"�������Ǻã��������õ���������"&Application("aqjh_baowuname")&"����һ���������"
            sql="update �û� set ����=false,��������=0,����='"& Application("aqjh_baowuname") &"' where ����='"&aqjh_name &"'"
        else
            mess=""& aqjh_name &"��Ȼ��ǰһ������Ȼ����ǰ�治Զ���ֽ�������"&Application("aqjh_baowuname")&"��"
            sql="update �û� set ��������=0,����='"& Application("aqjh_baowuname") &"' where ����='"&aqjh_name &"'"
	end if
end select
conn.execute ""&sql&""
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>���������</font>"&mess			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="#0000FF"
saycolor="#0000FF"
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
<title>�������</title></head>
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