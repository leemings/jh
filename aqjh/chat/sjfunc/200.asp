<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_jy then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҫ["&jhdj_jy&"]��,�ſ��Բ�����');}</script>"
	Response.End
end if
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
'�����뿪������Ѩ�ж�
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
says="<font color=green>�������񹦡�<font color=" & saycolor & ">"+cuan2(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'������
function cuan2(fn1,to1)
fn1=abs(fn1)
if fn1<10000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������书һ������10000�ģ���̫С����');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
cy=DateDiff("n",rs("����ʱ��"),now())
if cy<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-cy
	Response.Write "<script language=JavaScript>{alert('�������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if fn1>50000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����书��Ҫ����50000�ò���');}</script>"
	Response.End
end if
if fn1<10000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����书��ҪС��10000�ã�');}</script>"
	Response.End
end if
if rs("�书")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������书���㣡');}</script>"
	Response.End
end if
if rs("����")<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������100000');}</script>"
	Response.End
end if
if rs("����")<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������100000��');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ��ģ�������ʲô����');}</script>"
	Response.End
end if
if rs("֪��")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����֪�ʲ���1000��');}</script>"
	Response.End
end if
conn.execute "update �û� set �书=�书-" & fn1 & " where ����='" & aqjh_name &"'"
conn.execute "update �û� set ֪��=֪��-100 where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����-8000 where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����-15000 where ����='" &aqjh_name &"'"
conn.execute "update �û� set �书=�书+" & fn1*0.90 & " where ����='" & to1 &"'"
conn.execute "update �û� set �书��=�书��+20 where ����='" & to1 &"'"
cuan2= "##������<bgsound src=wav/dz.wav loop=1>������һ�񹦣������񹦣��벻������Ҫ�������˲ſ��Ե�<bgsound src=wav/dz.wav loop=1>���������ò����£�����Ϊ��������ױ������##���������ڰ�" & fn1 & "���书������%%�ˣ���;��ʧ10%�书��֪���½�100��%%��ָ�л��##����<font color=red>-8000</font>��,����<font color=red>-15000</font>!%%�书��������<font color=red>+20</font>��!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>