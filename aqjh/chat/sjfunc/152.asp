<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'Ѱ��ħ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")

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
says="<font color=green>��Ѱ��ħ����<font color=" & sayscolor & ">"+xunfaqi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)


function xunfaqi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ȼ�,w6,w9,����,����ʱ��,ְҵ,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
if DateDiff("s",rs("����ʱ��"),now())<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('Ϊ�˷�ֹ����ˢ��Ѱ��ħ����300���Ӳ���!');}</script>"
	Response.End 
end if
dj=rs("�ȼ�")
fla=rs("����")
w6w=rs("w9")
if fla<5000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��5000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ����ǳ���������ʲô��������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("ְҵ")<>"ð�ռ�" then
	Response.Write "<script language=JavaScript>{alert('���ð�ռң������ұ������ȥְҵת��Ϊð�ռң�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<50 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ50��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if


conn.execute "update �û� set ����=����-5000 where ����='" & aqjh_name &"'"
randomize()
r=int(rnd*25)+1
select case r
case 1
fq=add(w6w,"������",1)
conn.execute "update �û� set  w9='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##���ڼ���Ĵ�����͵����<font color=red>����������</font>�������ض�����."
case 2
fq=add(w6w,"����׶",1)
conn.execute "update �û� set  w9='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##�㴳����Ů����<font color=red>��������׶</font>�����������ߣ��������²�����."
case 3
fq=add(w6w,"Ѫ����",1)
conn.execute "update �û� set  w9='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##���ڽ������һ�������з����˱����˶�����<font color=red>����Ѫ����</font>."
case 4
fq=add(w6w,"������",1)
conn.execute "update �û� set  w9='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##���������ˣ��ѵ�һ����<font color=red>����������</font>���ܱ����ң������и�֮�˰�."
case 5
w6w=rs("w6")
fq=add(w6w,"�챦ʯ",1)
conn.execute "update �û� set  w6='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##�㷢��%%�ڴ���������,˳��һ��ԭ����һ��<font color=red>�챦ʯ</font>,���Ǽ��˿���ʯͷ����%%�Ŀڴ�,�Ѻ챦ʯ��������."
case 6
w6w=rs("w6")
fq=add(w6w,"�̱�ʯ",1)
conn.execute "update �û� set  w6='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##�㷢�ֿڴ����̹�����,˳��һ��ԭ����һ��<font color=red>�̱�ʯ</font>,���Ǽ��˿���ʯͷ����%%�Ŀڴ�,���̱�ʯ��������."
case 7
w6w=rs("w6")
fq=add(w6w,"����ʯ",1)
conn.execute "update �û� set  w6='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##�㷢�ֿڴ�����������,˳��һ��ԭ����һ��<font color=red>����ʯ</font>,���Ǽ��˿���ʯͷ����%%�Ŀڴ�,������ʯ��������."
case 8
fq=add(w6w,"ħ����ʯ",1)
conn.execute "update �û� set  w9='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"

xunfaqi="##���ڽ������һ��ǧ���������֦�Ϸ���һ��<font color=red>ħ����ʯ</font>��##����۹����ǼⰡ."
case 9
w6w=rs("w6")
fq=add(w6w,"���յ���",1)
conn.execute "update �û� set  w6='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##�����գ����Ϊ##����һ�����յ���."
case 10
fq=add(w6w,"���鵶",1)
conn.execute "update �û� set  w9='"&fq&"',����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##��ȥ����������,���ź���<font color=red>һ�Ѿ��鵶</font>."
case else
conn.execute "update �û� set ����ʱ��=now() where ����='"&aqjh_name&"'"
xunfaqi="##���������ǲ��ѽ��ʲô��û�ҵ�."
end select
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>