<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_jc then
	Response.Write "<script language=JavaScript>{alert('��ʾ����������Ҫ����ȼ�["&grade_jc&"]�ſ��Բ�����');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
'call dianzan(towho)
if Instr(Application("aqjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��" & towho & "�������������У����ܶ��䷢�ԡ�');parent.f2.document.af.towho.value='���';parent.f2.document.af.towho.text='���';parent.m.location.reload();</script>"
	Response.end
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>����������</font><font color=" & saycolor & ">"+shifang(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'������
function shifang(fn1)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ״̬ from �û� where ����='"& fn1 &"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����������û��������ڣ�');}</script>"
	Response.End
end if
if rs("״̬")<>"���" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('["& fn1 &"]��û�б������');}</script>"
	Response.End
end if
conn.execute "update �û� set ״̬='����',��¼=now(),�¼�ԭ��='"&aqjh_name&" ��������"&fn1&"' where ����='" & fn1 & "'"
shifang= "##:<font color=red>�ٸ����˿�"& fn1 &"���ĸĹ�����������,�����һ��ͬ�����һ�θĹ����µĻ��ᣬ�����ͷţ�</font>"
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & fn1 & "','������',now(),'���ĸĹ���ԭ������һ�Σ�')"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>