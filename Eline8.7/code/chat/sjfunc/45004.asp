<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'ħ��ˮ����wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_shqiu then
	Response.Write "<script language=JavaScript>{alert('ħ��ˮ������Ҫ�����ȼ�["&jhdj_shqiu&"]���Ĳſ��Բ�����');}</script>"
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
says="<font color=red>��ħ��ˮ����<font color=" & saycolor & ">"+shqiu(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'ħ��ˮ��
function shqiu(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w9,���� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
fla=rs("����")

if rs("����")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��1000�㰡��');window.close();}</script>"
	response.end
end if
duyao=rs("w9")
if isnull(duyao) or duyao=abate(rs("w9"),"ˮ����",1) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������һ��ˮ����û��Ҳ!');</script>"
	response.end
end if


rs.close
conn.execute "update �û� set ����ʱ��=now()  where ����='"&sjjh_name&"'"

randomize 
r=int(rnd*5)+1
select case r
case 1
shqiu=sjjh_name & "��<bgsound src=wav/Ye150.wav loop=1>�ڴ����ó�ˮ���򣬶���" & to1 & "���І����дʣ�ֻ��ˮ������õ�ɢ������â������<font color=red>" & to1 & "</font>�����ϣ�" & to1 & "���Ժ�����˯����." 
   rs.open "SELECT w9 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
   duyao=abate(rs("w9"),"ˮ����",1)
   conn.execute "update �û� set  ����=����-3000,w9='"&duyao&"' where ����='"&sjjh_name&"'"
   n=Year(date())
   y=Month(date())
   r=Day(date())
   s=Hour(time())
   f=Minute(time())
   m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
application("sjjh_dianxuename")=application("sjjh_dianxuename")& to1 & "|"&sj2&"|"&";"
rs.close	

case 2
shqiu=sjjh_name & "��<bgsound src=wav/Ye150.wav loop=1>�ڴ����ó�ˮ���򣬶���" & to1 & "���І����дʣ�ֻ��ˮ������õ�ɢ������â������<font color=red>" & to1 & "</font>�����ϣ������ˮ����ʧЧ." & to1 & "���úø����ˮ�������û���� " 
		conn.execute "update �û� set ����=����-1000 where ����='" & sjjh_name &"'"

case 3
shqiu=sjjh_name & "��<bgsound src=wav/Ye150.wav loop=1>�ڴ����ó�ˮ���򣬶���" & to1 & "���І����дʣ�ֻ��ˮ������õ�ɢ������â������<font color=red>" & to1 & "</font>�����ϣ������ˮ����ʧЧ." & to1 & "���úø����ˮ�������û���� " 
	conn.execute "update �û� set ����=����-1000 where ����='" & sjjh_name &"'"
	
case 4
shqiu=sjjh_name & "��<bgsound src=wav/Ye150.wav loop=1>�ڴ����ó�ˮ���򣬶���" & to1 & "���І����дʣ�ֻ��ˮ������õ�ɢ������â������<font color=red>" & to1 & "</font>�����ϣ������ˮ����ʧЧ." & to1 & "���úø����ˮ�������û���� " 
	conn.execute "update �û� set ����=����-1000 where ����='" & sjjh_name &"'"

case 5
shqiu=sjjh_name & "��<bgsound src=wav/Ye150.wav loop=1>�ڴ����ó�ˮ���򣬶���" & to1 & "���І����дʣ�ֻ��ˮ������õ�ɢ������â������<font color=red>" & to1 & "</font>�����ϣ������ˮ����ʧЧ." & to1 & "���úø����ˮ�������û���� " 
	conn.execute "update �û� set ����=����-1000 where ����='" & sjjh_name &"'"
	
case 6
shqiu=sjjh_name & "��<bgsound src=wav/Ye150.wav loop=1>�ڴ����ó�ˮ���򣬶���" & to1 & "���І����дʣ�ֻ��ˮ������õ�ɢ������â������<font color=red>" & to1 & "</font>�����ϣ�" & to1 & "���Ժ�����˯����." 
	 rs.open "SELECT w9 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
   duyao=abate(rs("w9"),"ˮ����",1)
   conn.execute "update �û� set  ����=����-5000,w9='"&duyao&"' where ����='"&sjjh_name&"'"
   conn.execute "update �û� set ״̬='��',��¼=now()+1,�¼�ԭ��='"&sjjh_name&" ��:"&fn1&"' where ����='" & to1 &"'"
   call boot(to1,"����������ߣ�"&sjjh_name&","&fn1)
   conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & to1 & "','��',now(),'" & fn1 & "')"
   rs.close	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>


