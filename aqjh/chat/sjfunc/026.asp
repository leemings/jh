<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'ɱ�����
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
'ɱ�����"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ʦ��,����,����,ɱ����,��ɱ��,����,��Ա�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("��Ա��")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[%%]��[<font color=red>##</font>]˵����������<font color=red>"&fn1&"</font>��𿨣�����ˣ�����һ��������"
	exit function
end if
if rs("����")<fn1*1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[<font color=red>%%</font>]��[<font color=red>##</font>]˵�����������<font color=red>"&fn1*1000&"</font>��������񶼿��������ˣ�����ô������أ�"
	exit function
end if
rs.close
rs.open "SELECT ����,grade,ʦ��,����,��ɱ��,��Ա��,֪��,����,ɱ���� FROM �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("ɱ����")>=3 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=red>[<b>�ٸ�</b>]��[##]<b>����</b>˵�������ɱ��<img src='picwords/3.gif'>�����ˣ����ܼ���ɱ������̫����~��û����ץ���������~��"
	exit function
end if
if rs("��ɱ��")<=0 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=" & saycolor & ">##��ʹ��[<b><font color=red>ɱ�����</font></b>]�����Լ�ɱ������##���ɱ������<font color=red>0</font>�ˣ��㻹Ҫ��ô��ѽ~�㿴[%%]��ȡЦ����~~�ٺ�~~"
	exit function
end if
if fn1>20 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=" & saycolor & ">##��ʹ��[<b><font color=red>ɱ�����</font></b>]�����Լ�ɱ�������������涨���������Ϊ<font color=red>20</font>Ŷ,�㿴[%%]��ȡЦ����~~�ٺ�~~"
	exit function
end if
if rs("��Ա��")>=fn1 then
    conn.execute("update �û� set ֪��=֪��+" & fn1 & ",����=����-" & fn1*1000 & ",��ɱ��=��ɱ��-" & fn1 & ",��Ա��=��Ա��-" & fn1 & " where ����='"&aqjh_name&"'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>��ɱ�������</font>[<font color=blue>##</font>]���Լ��Ʋ����ó���[<font color=red>" & fn1 & "</font>]��𿨣���ͨ����縺����ɱ����ʹ��ħ��˳���ļ�����[<font color=red>" & fn1 & "</font>]��ɱ��������֪������[<font color=red>" & fn1 & "</font>],���¼�����[<font color=red>" & fn1*1000 & "</font>]����������ѽ~"
	exit function
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
