<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'Ѱˮ�����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_xunshuijing then
	Response.Write "<script language=JavaScript>{alert('Ѱˮ������Ҫ�����ȼ�["&jhdj_xunshuijing&"]���Ĳſ��Բ�����');}</script>"
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
says="<font color=red>��Ѱˮ����<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'Ѱˮ����
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
fla=rs("����")
if rs("����")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��1000�㰡��');window.close();}</script>"
	response.end
end if
rs.close
conn.execute "update �û� set ����ʱ��=now()  where ����='"&sjjh_name&"'"
randomize 
r=int(rnd*9)+1
select case r
case 1
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"
case 2
xunshuijing=sjjh_name & "��ǧ�����Ѱ��<font color=red>ˮ����</font>�����䣬ȴ������һ��ˮ���ﱻ���ҵ��ˣ��Ͻ�ϴϴˮ�����ھ�����ħ����."
	rs.open "SELECT w9 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w9"),"ˮ����",1)
	conn.execute "update �û� set  w9='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close

case 3
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"
	
case 4
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"

case 5
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"

case 6
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"

case 7
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"

case 8
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"

case 9
xunshuijing=sjjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲôˮ����," & sjjh_name & "���100�㷨��!" 
	conn.execute "update �û� set ����=����-100 where ����='" & sjjh_name &"'"
	
case 10
xunshuijing=sjjh_name & "��ǧ�����Ѱ��<font color=red>ˮ����</font>�����䣬ȴ������һ��ˮ���ﱻ���ҵ��ˣ��Ͻ�ϴϴˮ�����ھ�����ħ����."
	rs.open "SELECT w9 FROM �û� WHERE ����='"&sjjh_name&"'",conn
	duyao=add(rs("w9"),"ˮ����",1)
	conn.execute "update �û� set  w9='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
