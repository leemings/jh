<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_name<>Application("aqjh_mengzhu") then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻��������Ϲ��ʲô�Ұ���');}</script>"
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
says=mzl(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'������
function mzl(fn1,to1)
saysok="<img src=pic/mengzhulin.gif><font color=ff0013>�������</font>"
'�ж϶Է�
if to1=aqjh_name or to1="���" or to1=Application("aqjh_automanname") then
 	Response.Write "<script language=JavaScript>{alert('���棺������˭��');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='"&to1&"'",conn,1,3
to_grade=rs("grade")
to_tj=rs("ͨ��")
to_yin=rs("����")
rs.close
if to_grade>5 then
 mzl=saysok&"##����%%ʧ�ܣ���Ϊ%%�ǹٸ�����"
 exit function
end if
fn1=trim(fn1)
if fn1="����" or fn1="" then
	Response.Write "<script language=JavaScript>{alert('���������:�������������ͷ���ն�ף�����[�������ã�����ɾ��id]');}</script>"
	Response.End
elseif fn1="����" then
 if to_tj=true then
  mzl=saysok&"��������##��%%���н���ʧ�ܣ��Է�����ͨ���а���" 
  exit function
 end if
 rs.open "select ����ʱ�� FROM �û� WHERE ����='"&aqjh_name&"'",conn,1,3
 sj=DateDiff("s",rs("����ʱ��"),now())
 if sj<55 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=55-sj
	Response.Write "<script language=JavaScript>{alert('�������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
 end if
 randomize()
 rnd1=int(rnd*400000)+100000
 conn.execute "update �û� set ����=����+" & rnd1 & " where ����='" & to1 & "'"
 conn.execute "update �û� set ����ʱ��=now() where ����='" &aqjh_name& "'"
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','������������"&rnd1&"��')"
 mzl=saysok&"��%%�Խ����й�����������##�ؽ���%%"&rnd1&"������" 
elseif fn1="�ͷ�" then
 if to_yin<100000 then
  mzl=saysok&"��������##��%%���з���ʧ�ܣ�%%�������⵰�ˣ����ȷŹ�%%�ɣ�" 
  exit function
 end if
 randomize()
 rnd1=int(rnd*90000)+10000
 conn.execute "update �û� set ����=����-" & rnd1 & " where ����='" & to1 & "'"
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','�����������"&rnd1&"��')"
 mzl=saysok&"%%̫�����ˣ���������##��%%����"&rnd1&"������" 
elseif fn1="ն��" then
 conn.execute "update �û� set ״̬='��',�¼�ԭ��='����"&aqjh_name&"ն��' where ����='" & to1 &"'"
 mzl= saysok&"��������##����:�˷�%%ն����!"
 call boot(to1,"ն�����������ߣ�"&aqjh_name&","&"��������")
 conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','ն��',now(),'����ն��')"
elseif fn1="����" then
 call boot(to1,"���ˣ������ߣ�"&aqjh_name&",��������")
 mzl= saysok&"<bgsound src=wav/gf.wav loop=1>һ�����%%�γ���������(����=##)"
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','��������')"
else
mzl="<bgsound src=wav/luo.wav loop=3><table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chat/f2/img/rightct3.gif></td><td background=../chat/f2/img/rightct4.gif></td><td><img  src=../chat/f2/img/rightct1.gif></td></tr><tr><td background=../chat/f2/img/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr><td valign=center align=center><font style='font-size:9pt;color:red'>========��������========</font></td></tr><tr><td valign=center align=center><font style='font-size:9pt;color:green'>"&fn1&"(##)</td></tr></table></td><td background=../chat/f2/img/rightct08.gif></td></tr><tr><td><img src=../chat/f2/img/rightct9.gif></td><td background=../chat/f2/img/rightct10.gif></td><td><img src=../chat/f2/img/rightct11.gif></td></tr></table></td></table>"
end if
end function
%>