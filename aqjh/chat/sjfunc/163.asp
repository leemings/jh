<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'Ѱ�ҷ���
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��Ѱ�ҷ�����<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'Ѱ�ҷ���
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
if DateDiff("s",rs("����ʱ��"),now())<40 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('Ϊ�˷�ֹ����ˢ��Ѱ�ҷ�����40���Ӳ���!');}</script>"
	Response.End 
end if
dj=rs("�ȼ�")
fla=rs("����")
if rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��5000�㰡��');window.close();}</script>"
	response.end
end if
if rs("ְҵ")<>"ð�ռ�" then
	Response.Write "<script language=JavaScript>{alert('���ð�ռң������ұ������ȥְҵת��Ϊð�ռң�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲��ܲ�����');}</script>"
	Response.End
end if
if dj<80 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ80��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

rs.close
conn.execute "update �û� set ����=����-5000,����ʱ��=now()  where ����='"&aqjh_name&"'"
randomize 
r=int(rnd*9)+1
select case r
case 1
xunfaqi=aqjh_name & "�㴳��" & to1 & "�����<font color=red>������������</font>�����������ߣ��������²�����."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"��������",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close


case 2
xunfaqi=aqjh_name & "����" & to1 & "����Ĵ�����͵����<font color=red>������</font>�����ض�����."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"������",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close

	
case 3
xunfaqi=aqjh_name & "���ڽ������һ�������з����˱����˶�����<font color=red>������ڤ��</font>."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"��ڤ��",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "���������ˣ��ѵ�һ����<font color=red>������ȡ��</font>���ܱ����ң������и�֮�˰�."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"��ȡ��",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "�㷢��" & to1 & "�ڴ���������,˳��һ��ԭ����һ��<font color=red>����ʯ</font>���Ǽ��˿���ʯͷ����" & to1 & "�Ŀڴ�,�Ѻ���ʯ��������.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"����ʯ",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close



case 6
xunfaqi=aqjh_name & "�㷢��" & to1 & "�ڴ���������,˳��һ��ԭ����һ��<font color=red>����ʯ</font>���Ǽ��˿���ʯͷ����" & to1 & "�Ŀڴ�,������ʯ��������.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"����ʯ",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close

case 7
xunfaqi=aqjh_name & "�㷢��" & to1 & "�ڴ���������,˳��һ��ԭ����һ��<font color=red>����ʯ</font>���Ǽ��˿���ʯͷ����" & to1 & "�Ŀڴ�,������ʯ��������.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"����ʯ",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close

case 8
xunfaqi=aqjh_name & "���ڽ������һ��ǧ���������֦�Ϸ���ʧ���Ѿõ�<font color=red>������</font>" & aqjh_name & "�Ȿ�����������ޱȿ������Ϊ���ֵ�һ����"
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"������",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "�㷢��" & to1 & "�ڴ�����һ������Ȩ�����̴�" & to1 & "�ڴ�͵����<font color=red>����ȭ��</font>�����ܻؼҰ�����ȭ����.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"����ȭ",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "�㷢��" & to1 & "�ڴ�����һ�����������̴�" & to1 & "�ڴ�͵����<font color=red>������</font>�����ܻؼҰ�����ȭ����.."
	rs.open "SELECT w10 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	duyao=add(rs("w10"),"ʥ����",1)
	conn.execute "update �û� set  w10='"&duyao&"' where ����='"&aqjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
ng
end function
%>
