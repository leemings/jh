<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������ʲôѽ����Ĺ���ȼ��ɲ���ѽ��');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��������<font color=" & saycolor & ">"+jiangli(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function jiangli(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,grade,��� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
menpai=rs("����")
if rs("grade")<4 and (rs("���")<>"����" or rs("���")<>"����") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ʲôѽ����Ĺ���ȼ��ɲ���ѽ��');}</script>"
	Response.End
end if
fn1=int(abs(fn1))
if fn1>100000 or fn1<100  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ʲôѽ����Ĺ���ȼ��ɲ���ѽ��');}</script>"
	Response.End
end if
rs.close
rs.open "select h FROM p WHERE a='" & menpai & "'",conn,2,2
if rs("h")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ֻ�У�"&rs("h")&"��,��������ô���Ǯѽ��');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,���ɻ���,�ȼ� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if trim(rs("����"))<> trim(menpai) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������˰ɣ�["&to1&"]���������ɵ��ӣ�');}</script>"
	Response.End
end if
if rs("���ɻ���")<-(rs("�ȼ�")*20000) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ѽ��["&to1&"]�Ѿ�����̫���Ǯ��');}</script>"
	Response.End
end if
conn.execute "update p set h=h-" & fn1 & " where a='" & menpai & "'"
conn.execute "update �û� set ����=����+" & fn1 & ",���ɻ���=���ɻ���-"& fn1 &" where ����='" & to1 &"'"
jiangli="##���Լ�����["&menpai&"]�Ļ��𷢸�����%%,"& fn1 & "����Ϊ������%%�ֵ�ֱ��,����˵лл!"
fn1="���ɣ�"&menpai&"  ������"&fn1
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','����','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>