<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�ͷ�ͽ��
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
'�ͷ�ͽ��"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ʦ��,����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<2000 then
	Response.Write "<script language=JavaScript>{alert('���Լ�û�����ˣ���֪����ܲ����㻹��ʲô�ʸ�ͷ���ͽ�ܣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "SELECT ����,grade,ʦ��,���� FROM �û� where ����='" & to1 & "' and ʦ��='"&aqjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[�ٸ�]��[##]˵��[%%]�������ͽ�ܰ����㲻�ܳͷ�[%%]��������ȥ����Ŷ������Բ��˶����ߣ�"
	exit function
end if
if fn1>10000 then
    conn.execute("update �û� set ����=����-2000 where ����='" & aqjh_name & "'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="##���ͽ��%%�ͷ���" & fn1 & "��������������ʦ��##̫����,�Լ�������2000����£����ѵ���Ϊ��ʦ����"
	exit function
end if
if rs("����")>=fn1 then
    conn.execute("update �û� set ����=����-" & fn1 & " where ����='" & to1 & "'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=green>���ͷ�ͽ�ܡ�</font>##����ͽ��</font>%%<font color=red>������˵:�����ͽ�����Ҳ�����,����Σ��ʦ�ţ�Σ���������������ҷ���</font>" & fn1 & "<font color=red>����������û��ϲ����������!</font>"
	exit function
end if
fakuan="<font color=red>##</font><font color=red>�㻹�ǾͷŹ���ͽ��</font><font color=red>%%</font><font color=red>�ɣ����Ѿ���û��Ե�ˣ���" & fn1&"��������û��������������ӡ��ò�����</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>