<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�չ�����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=fakuan(mid(says,i+1)+0,towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�չ�����"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ����='" & to1 & "'",conn,2,2
if rs("������")>50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>���չ����ˡ�</font><font color=blue>##</font>�������˰ɣ�<font color=blue>%%</font>���Ѿ������չ�һ���ˣ����������ˣ�</font>"
	exit function
end if
if rs("����")>1000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>���չ����ˡ�</font><font color=blue>##</font>�������˰ɣ�<font color=blue>%%</font>�����������˰���</font>"
	exit function
end if
if rs("�ȼ�")>18 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>���չ����ˡ�</font><font color=blue>##</font>�������˰ɣ�<font color=blue>%%</font>�����������˰���</font>"
	exit function
end if
if fn1>5000 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=" & saycolor & "><font color=red>���չ����ˡ�</font><font color=blue>##</font>�չ˿��ֽ������ˣ������õ���ɮ�����ִ���˴��ڣ�##Ҫ����<font color=red>" & fn1 & "</font>��������<font color=blue>[%%]</font>����������Ŀ̫�࣬ʧ�ܣ���λ���˴�Ӧ�����ֻ��<font color=red>5000</font>�㣡</font>"
	exit function
end if
if rs("����")>=fn1 then
    conn.execute("update �û� set ������=������+100000,����=����+1000000,����=����+" & fn1 & " where ����='" & to1 & "'")
    conn.execute("update �û� set ����=����+1000 where ����='" & aqjh_name & "'")
  	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>���չ����ˡ�</font><font color=blue>[##]</font>���տ������磬��վ��֮����������<font color=blue>[%%]</font>��<font color=blue>[%%]</font>��������<font color=red>100��</font>��������������<font color=red>10��</font>������������<font color=red>" & fn1 & "</font>��<font color=blue>[##]</font>���ڰ������ˣ���������<font color=red>1000</font>�㣬������һЩĪ����������ԣ�</font>"
	exit function
end if
fakuan="<font color=blue><font color=red>##</font>,<font color=red>%%</font>���Ѿ���<font color=red>"&fn1&"</font>����������������</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
