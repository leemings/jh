<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="../../const3.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####���䴦��#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(8)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ�ÿ�Ƭ��');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
f=Minute(time())
aqjh_kptime=aqjh_chatkptime
if f>aqjh_kptime and chatinfo(0)<>aqjh_chat1 then
 Response.Write "<script language=javascript>{alert('��ʾ���Ƕ�Թ�����ÿ�����ÿСʱ��ǰ"&aqjh_kptime&"��');}</script>"
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
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=trim(mid(says,i+1))
if fnn1<>"������" then call dianzan(towho) 
says=kapian(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��Ƭ
function kapian(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if aqjh_name=to1 and instr("���ҿ�;������;���˿�;���߿�;���׿�;���˿�;��˰��;����;˥��;����;���ֿ�;������;���޿�;������;�����;�ݺ���;�޻���;�Զ���;������;�氮��;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�������Լ����в�����');</script>"
	Response.End
end if
if to1="���" and instr("���ҿ�;������;���˿�;���߿�;���׿�;���˿�;��˰��;����;˥��;����;���ֿ�;������;���޿�;������;�����;�ݺ���;�޻���;�Զ���;������;�氮��;���ӿ�;�ֻ���;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�����ܴ�ҽ��в�����');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w5,���� from �û� where  ����='"&aqjh_name&"'",conn,2,2
mycard=abate(rs("w5"),fn1,1)
select case fn1
case "����"
 rs.close
 rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("sl")<>"" and rs("sl")<>"��" then
	kapian="<font color=blue>����Ƭ��<font color=" & saycolor & ">##�����Ѿ������鸽���ˣ�����ʹ��["&fn1&"]�����Ȱ����ϵ��������߰ɣ�</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�������鸳���ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ����["&fn1&"]���񰡣�������ͬ��...</font>"
case "��Ա��"
 rs.close
 rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("�ȼ�")<90 and rs("ת��")<1 then
	kapian="<font color=blue>��ħ����Ƭ��<font color=" & saycolor & ">##����ĵȼ�̫���ˣ�����90��������ʹ��["&fn1&"]������ݵ�ɣ�</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
 if rs("��Ա�ȼ�")>2 then
	kapian="<font color=blue>��ħ����Ƭ��<font color=" & saycolor & ">##���㲻����ѻ�Ա������ʹ��["&fn1&"]������ϵվ���ɣ�</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
 conn.Execute ("update �û� set ��Ա�ȼ�=2,��Ա����=��Ա����+60 where ����='"&aqjh_name&"'")
 kapian="<font color=green>��ħ����Ƭ��<font color=" & saycolor & ">##����90����ʹ����["&fn1&"]����Ա�ȼ���Ϊ2������Ա��������2���£�����ˬ������</font>"
case "����"
 rs.close
 rs.open "select allvalue,��Ա�ȼ�,�ȼ�,ת��,sl from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("�ȼ�")>90 or rs("ת��")>0 then
		Response.Write "<script language=javascript>alert('��ʾ��90�����ϻ�ת���˻�2�����ϻ�Ա����ʹ�ñ��񿨣�');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
 if rs("sl")<>"" and rs("sl")<>"��" then
	kapian="<font color=blue>����Ƭ��<font color=" & saycolor & ">##�����Ѿ������鸽���ˣ�����ʹ��["&fn1&"]�����Ȱ����ϵ��������߰ɣ�</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update �û� set sl='����',slsj=now()+1 where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�������鸳���ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ����["&fn1&"]���񰡣�������ͬ�ڣ�Ը���ݵ��ٶ�����3.5��...</font>"
case "��ɱ��"
 rs.close
 rs.open "select allvalue,��Ա�ȼ�,�ȼ�,ת��,ɱ���� from �û� where ����='"&to1&"'",conn,2,2
 if rs("ɱ����")<=2 then
		kapian="<font color=red><bgsound src=wav/003.mid loop=1>��ħ����Ƭ��<font color=" & saycolor & ">"&aqjh_name&","&to1&"����ɱ����û��<img src='picwords/3.gif'>�ˣ����ܶ���ʹ�÷�ɱ��Ŷ~!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update �û� set ɱ����=5,��ɱ��=��ɱ��+10 where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�����ڶԷ��Ѿ�����ɱ�ˣ�');}</script>"
	kapian="<font color=red><b>��ħ����Ƭ��</b><font color=" & saycolor & ">##͵͵��%%ʹ��["&fn1&"]��ʹ��%%������Ҳɱ�������ˣ���Ϊ%%������ɱ��ħ�ĳƺţ�ɱ�˼�¼����<img src='picwords/1.gif'><img src='picwords/0.gif'>��...</font>"
case "����"
 rs.close
 rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("sl")<>"" and rs("sl")<>"��" then
	kapian="<font color=blue>����Ƭ��<font color=" & saycolor & ">##�����Ѿ������鸽���ˣ�����ʹ��["&fn1&"]�����Ȱ����ϵ��������߰ɣ�</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�������鸳���ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ����["&fn1&"]���񰡣�������ͬ��...</font>"
case "����"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]����������["&to1&"]�����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]���񰡣�����ͬ��...</font>"
case "˥��"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update �û� set sl='˥��',slsj=now()+3 where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]����������["&to1&"]�����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]���񰡣�����ͬ��...</font>"
case "����"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]����������["&to1&"]�����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]���񰡣�����ͬ��...</font>"
case "����"
	conn.Execute ("update �û� set sl='��',slsj=now() where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�ɹ���������鸳����');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ����["&fn1&"]���񰡣�88��see tomorrow!</font>"
case "������"
 rs.close
 rs.open "select �Ա�,�ȼ� from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("�ȼ�")<18 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�̫���ˣ�');}</script>"
	response.end
 end if
 my_xb=rs("�Ա�")
 rs.close
 rs.open "select �Ա� from �û� where ����='"&to1&"'",conn,2,2
 to_xb=rs("�Ա�")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('����������Ҳ�����ǲ���ͬ��������');}</script>"
	response.end
 end if
 conn.Execute ("update �û� set ����=����-100,����=����+500,����=����+500 where ����='"&to1&"'")
 conn.Execute ("update �û� set ����=����-100,����=����+500,����=����-100,allvalue=allvalue+30 where ����='"&aqjh_name&"'")
 kapian="<bgsound src=wav/baobao.wav loop=1><font color=green>����Ƭ��<font color=" & saycolor & ">##��%%ʹ���˱�������������Ը�Գ�����%%�ڽ����Ĵ�������<img src=card/baobao.gif>�������������Ľ����ͬʱ##�ľ���ֵ������30��,�㿴����##�ֵ�������ҹ˯���ž�����У�<font color=red><b>����������˧����Ů������!</b></font>"
case "�氮��"
 rs.close
 rs.open "select �Ա�,�ȼ� from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("�ȼ�")<18 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�̫���ˣ�');}</script>"
	response.end
 end if
 my_xb=rs("�Ա�")
 rs.close
 rs.open "select �Ա� from �û� where ����='"&to1&"'",conn,2,2
 to_xb=rs("�Ա�")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('���������㰮ʲô�����ǲ���ͬ��������');}</script>"
	response.end
 end if
 conn.Execute ("update �û� set ����=����+50,����=����+50,allvalue=allvalue+50 where ����='"&to1&"'")
 conn.Execute ("update �û� set ����=����+50,����=����+50,allvalue=allvalue-50 where ����='"&aqjh_name&"'")
 kapian="<bgsound src=wav/zhenai.wav loop=1><font color=green>����Ƭ��<font color=" & saycolor & ">##��%%ʹ�����氮����������Ը�Գ�����%%�ڽ����Ĵ�������<img src=card/zhenai.gif>���氮��%%50����,������˽�ķ��ף���н�����ʿ�ۺ찡��ͬʱ˫���ĵ��º�����������50�㣬�㿴����##�ֵ��������ˣ���У�<font color=red><b>�±����һ�Ҫ�㰮�ҿ�����</b></font>"
case "����"
	dim sl(4)
	sl(0)="����"
	sl(1)="����"
	sl(2)="˥��"
	sl(3)="����"
	sl(4)="����"
	randomize timer
	sss=int(rnd*4)+1
	if sss=5 then sss=4
	conn.Execute ("update �û� set sl='"&sl(sss)&"',slsj=now()+3 where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]���ڵõ�����["&sl(sss)&"]������');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ����["&fn1&"]���񰡣�������ͬ��...��ô���["&sl(sss)&"]���Լ������ϡ�</font>"
	erase sl
case "�����"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update �û� set ����=false,����ʱ��=now() where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ������["&to1&"]������״̬��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ���˽����Ƭ��Ҳ��֪����һλ��ϺҪ��ù��...</font>"
case "�ݺ���"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update �û� set ����=int(����/3),����=int(����/3) where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]����������ֻʣԭ����1/3�ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##�����̲����̣��ó��Լ��ݺ�������%%��ͷ�ϴ�ȥ,%%ֻ����ǰһ�ڣ�������������ʧ���..</font>"
case "��Ѫ��"
	conn.Execute ("update �û� set ����=����+50000,����=����+50000  where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]��������������������5��㣡');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ���˲�Ѫ������һ�¿����и��ˣ���������������5��㣡</font>"
case "��Ǯ��"
	conn.Execute ("update �û� set ���=���+88880000  where  ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]�Ĵ��������8888��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ������Ǯ�����Լ���С�ɰ���װ�����ˣ�8888��ѽ��</font>"
case "���㿨"
 rs.close
 rs.open "select �ȼ�,ת��,���� from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("����")="����" or rs("����")="����" or rs("����")="����" or rs("�ȼ�")>90 or rs("ת��")>0 then
		Response.Write "<script language=javascript>alert('��ʾ���㲻�ǽ������ɵ��˻�������ת���ˣ��ȼ���90�����ˣ�����ʹ�����ֿ�Ƭ��');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	conn.Execute ("update �û� set allvalue=allvalue+1000  where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]ʹ�������㿨���ݵ�������1000�㣡');}</script>"
	kapian="<font color=red>�����ɿ�Ƭ��<font color=" & saycolor & ">##ʹ�������㿨���ǡ��������ݵ�Ҳ�ӵ㣬�����и���������������1000�㣡</font>"
case "�ӽ�"
 rs.close
 rs.open "select �ȼ�,ת��,����,��Ա��,���,grade from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("����")="����" or rs("��Ա��")>1000 or rs("����")<>"�ٸ�" or rs("���")>2000 or rs("����")="����" or rs("����")="����" then
		Response.Write "<script language=javascript>alert('��ʾ����Ľ𿨽��̫������㲻�ǹٸ����ˣ����Բ���ʹ�����ֿ�Ƭ��');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	conn.Execute ("update �û� set ��Ա��=��Ա��+100  where  ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]�Ľ�������100����');}</script>"
	kapian="<font color=red>�����ɿ�Ƭ��<font color=" & saycolor & ">##ʹ��������ר�п�Ƭ�ӽ𿨣��Լ��Ľ�ⶼװ�����ˣ�100����ѽ��</font>"
case "��Ǯ��"
 rs.close
 rs.open "select �ȼ�,ת��,����,��� from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("����")="����" or rs("����")="����" or rs("����")="����" or rs("���")>500000000 then
		Response.Write "<script language=javascript>alert('��ʾ���㲻�ǽ������ɻ������Ǯ��5�ڣ�����ʹ�����ֿ�Ƭ��');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	conn.Execute ("update �û� set ���=���+60000000  where  ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]�Ĵ��������6000��');}</script>"
	kapian="<font color=red>�����ɿ�Ƭ��<font color=" & saycolor & ">##ʹ�������ɿ�Ƭ��Ǯ�����Լ���С�ɰ���װ�����ˣ�6000��ѽ��</font>"
case "������"
	conn.Execute ("update �û� set �书=�书+10000  where  ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]ʹ�������������书����1��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ�������������书���Ǵ�������ǣ�����������Ҫ��̫ƽ�ˣ�</font>"
case "�䱦��"
	conn.Execute ("update �û� set ����=false,��������=0,����='"& Application("aqjh_baowuname") &"' where ����='"&aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('��ʹ�����䱦������ÿ�������1ö��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##�������Ǻã���Ȼʹ���˴�˵�е��䱦��������˽�������"&Application("aqjh_baowuname")&"<img src=img/z1.gif>1ö����һ�������������</font>"
case "������"
	conn.Execute ("update �û� set ����=����+300,����=����+300  where  ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]ʹ���˹������������ͷ�������300�㣡');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ���˹������������ͷ�����������ǣ����������ֶ��һλ�����ˣ�</font>"
case "�ӵ㿨"
	conn.Execute ("update �û� set allvalue=allvalue+5000  where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]ʹ���˼ӵ㿨���ݵ�������5000�㣡');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ���˼ӵ㿨���ǡ��������ݵ�Ҳ�ӵ㣬�����и�����</font>"
case "������"
%>
<!--#include file="wnkif.asp"-->
<%
 if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))=0 then
  kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & sayscolor & ">"&aqjh_name&"��"&to1&"ʹ���˹�����,��ϧ{"&to1&"}������������״̬!�װ��˷���һ�Ź�������</font>"
 else
  rs.close
  rs.open "select id,��Ա�ȼ�,�ȼ�,����,����,����ͷ��,�Ա�,��������,��ż,ͨ��,ת�� from �û� where ����='" & to1 &"'",conn,2,2
  jhid=rs("id")
  hydj=rs("��Ա�ȼ�")
  jhdj=rs("�ȼ�")
  jhsf=rs("����")
if Instr(Application("aqjh_guibin"),"|" & to1 & "|")<>0 then
 jhsf="���"
end if
if Instr(Application("aqjh_admin_send"),"|" & to1 & "|")<>0 then 
 jhsf="����"
end if
if rs("��ż")=Application("aqjh_user") and rs("�Ա�")="Ů" then
 jhsf="վ������"
end if
if Application("aqjh_mengzhu")=to1 then 
 jhsf="��������"
end if
  if rs("ͨ��")=True then
   jhmp="ͨ����"
  else
   jhmp=rs("����")
  end if
  jhtx=rs("����ͷ��")
  sex=rs("�Ա�")
   myzs=rs("ת��")
   mypeiou=rs("��ż")
'  nowinroom=session("nowinroom")
  Application.Lock
  onlinelist=Application("aqjh_onlinelist"&nowinroom)
  onlinenum=UBound(onlinelist)
  for i=1 to onlinenum step 1
   onlinexx=split(onlinelist(i),"|")
   if onlinexx(0)=to1 then
   aqjh_zm=to1&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&jhdj&"|"&jhid&"|"&hydj&"|0"&"|"&onlinexx(9)&"|"&mypeiou&"|"&myzs
   onlinelist(i)=aqjh_zm
   exit for
   end if
  next
  Application("aqjh_onlinelist"&nowinroom)=onlinelist
  aqjh_zanli=split(application("aqjh_zanli"),";")
  for x=0 to UBound(aqjh_zanli)
   dxname=split(aqjh_zanli(x),"|")
   if dxname(0)="!"&to1&"!" then
    dxcz=aqjh_zanli(x)&";"
    application("aqjh_zanli")=replace(application("aqjh_zanli"),dxcz,"")
    kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & sayscolor & ">"&aqjh_name&"��"&to1&"ʹ���˹�����,�ɹ��ذ�{"&to1&"}������״̬�ϻ���!</font>"
    exit for
   end if
  next
  Application.UnLock
 end if
case "�Զ���"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update �û� set ����ʱ��=now()  where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]ʹ���˳Զ������������ٱ����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʵ�ڶ�%%����Ϊ������ȥ��ʹ���˳Զ�����%%���һ����������ȥ������û����...</font>"
case "������"
	conn.Execute ("update �û� set ����ʱ��=now()  where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('��["&aqjh_name&"]ʹ���˱����������ڱ����ɹ��ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##:����ţ��ұ����ұ�....�ӿڴ����ó��������������ɹ���</font>"
case "���ҿ�"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update �û� set �书=int(�书/3)  where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]ʹ���˵��ҿ������书ֻʣ1/3�ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##:����ţ�˭Ҳ��Ҫ�����ң���ҪΪ�������ʹ���˵��ҿ���"&to1&"���书ʧȥ���....</font>"
case "���޿�"
%>
<!--#include file="wnkif.asp"-->
<%
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
        if rs("����")="�ٸ�" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�ù��޿�!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
        if rs("����")="����" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶ����ɵ�������<font color=red>[["&to1&"]]</font>ʹ�ù��޿�!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
	conn.Execute ("update �û� set ����='��ʬ��',״̬='��ʬ��' where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]ʹ���˹��޿������Ѿ�����˽�ʬ�ˣ�');}</script>"
	kapian="<img src=card/guaishou.jpg><font color=green>����Ƭ��<font color=" & saycolor & ">##:����ţ�"&to1&"����˽�ʬ������С��Ŷ������µĻ���������ߣ��ҿ��Ա������....</font>"
case "�崿��"
	conn.Execute ("update �û� set ������=0  where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('��ʹ�����崿������������0��ɣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##�û����ȥ��ʱ��ѽ����</font>"
case "���׿�"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" &saycolor& ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&aqjh_name&"�����ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from �û� where ����='"&to1&"'",conn
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&aqjh_name&"���˼��ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing

    end if
    if rs("�Ա�")=sex then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" &saycolor& ">"&aqjh_name&"��"&to1&"ʹ�����׿�ʧ��:ԭ����"&to1&"��"&aqjh_name&"��ͬ��!û�����˲���Ҳ�밡��</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.Execute ("update �û� set allvalue=allvalue+50  where ����='" &aqjh_name&"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]ʹ���˼ӵ㿨���ݵ�������50�㣡�˿��ɻ��׵���������');}</script>"
    kapian="<table border=0><tr><td><img src=card/ai1.gif></td><td><font color=green><bgsound src=123.wav loop=1>����Ƭ��<font color=" &saycolor& ">"&aqjh_name&"��"&to1&"ʹ�����׿�,������Ը�Գ�����"&to1&"�ڽ����Ĵ�������<img src=card/ai.gif>����!�������Ľ����ͬʱ"&aqjh_name&"���ݵ�����50�㣡�㿴����"&aqjh_name&"�ֵö��������ˣ���У�<img src=card/qin.gif></font></td></tr></table>" 
case "���Կ�"
	rs.close
	rs.open "select �Ա�,��ż FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
	sex=rs("�Ա�")
	pl=rs("��ż")
	rs.close
    if pl="��" then 
        if sex="��" then
               sql="update �û� set �Ա�='Ů' WHERE ����='" & aqjh_name & "'"
               xb="Ư����Ů��"
         end if 
          if sex="Ů" then 
              sql="update �û� set �Ա�='��' WHERE ����='" & aqjh_name & "'"
              xb="Ӣ�����к�"
          end if
          Set Rs=conn.Execute(sql)  
            bianxi=aqjh_name&"ʹ�ñ��Կ���,������Ը�Գ�,�����"&xb&"!" 
        else
          bianxi="ʹ�ñ��Կ�ʧ��!ԭ��:"&aqjh_name&"���м��ҵ�����,��ô�������ѽ!��������������!Ϊ�˳ͽ���"&aqjh_name&"�����ֵ��°ܻ�����,�ش�û��"&aqjh_name&"�ı��Կ�!"
        end if
	kapian="<img src=card/bxk.jpg></td><td><font color=green>����Ƭ��<font color=" & saycolor & ">##������ͷ����¡�������Ƚ�ѽ���ұ䣬�ұ䡭��(�����ڱ�ɿ��նۣ�����������������<br>�������"&bianxi&"</font>"
case "��鿨"
	rs.close
	rs.open "select ��ż,boy FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
	peiou=rs("��ż")
	myboy=rs("boy")
	if isnull(myboy) or myboy="" then
		myboy=""
	else
		zt=split(myboy,"|")
		if UBound(zt)<>7 then
			myboy=""
		end if
	end if
	if peiou="��" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��"&aqjh_name&"������Ǯѽ������û����żѽ��');</script>"
		Response.End
	end if
if myboy<>"" then
 conn.execute "insert into gry(boy,fm1,fm2) values ('"&myboy&"','"&aqjh_name&"','"&peiou&"')"
end if
	conn.Execute ("update �û� set ��ż='��',boy='',boysex='' where ����='" & aqjh_name &"'")
	rs.close
	rs.open "select ��ż FROM �û� WHERE ����='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update �û� set ��ż='��',boy='',boysex='' where ����='"&peiou&"'")
	end if
	rs.close
	kapian="<img src=card/lifen.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">##����ǰ��󣬾����Լ�һ��˼�붷����ʹ������鿨,�����������["&peiou&"]����ˡ���</font>"
if myboy<>"" then kapian=kapian&"<br>���¶�Ժ��˫����С���Ѿ��͵��¶�Ժ���������޹��İ�������"
case "���׿�"
%>
<!--#include file="wnkif.asp"-->
<%
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&aqjh_name&"�����ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from �û� where ����='"&to1&"'",conn,2,2
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&aqjh_name&"���˼��ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("�Ա�")
    if rs("��ż")<>"��" then
	    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��"&to1&"ʹ�����׿�ʧ��:ԭ����"&to1&"���м��ҵ���</font>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("����")="�ٸ�" then 
           kapian="�����ԶԹ���Աʹ�����׿�����"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,2,2
    if rs("��ż")<>"��" then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��"&to1&"ʹ�����׿�ʧ��:ԭ�����Լ��Ѿ����м��ҵ�����!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("�Ա�")=sex then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��"&to1&"ʹ�����׿�ʧ��:ԭ����"&to1&"��"&aqjh_name&"��ͬ��!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update �û� set ��ż='"&aqjh_name&"' where ����='"&to1&"'"
    conn.execute "update �û� set ��ż='"&to1&"' where ����='"&aqjh_name&"'"
    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��"&to1&"ʹ�����׿�,������Ը�Գ�����"&to1&"��Ϊ��!</font>"
case "������"
%>
<!--#include file="wnkif.asp"-->
<%
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
        if rs("����")="�ٸ�" then 
            kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�ô�����!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"�ó����������Ĵ�����,��"&to1&"��ץ��!"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update �û� set ״̬='��',��¼=now()+3 where ����='" & to1 & "'"
            call boot(to1,to1&"��"&aqjh_name&"ʹ���˴�����")
        else
           kapian=kapian&mzky
        end if
case "���˿�"
%>
<!--#include file="wnkif.asp"-->
<%
      rs.close
      rs.open "select �ȼ�,ת��,����,���,��¼ FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
      dlsj=DateDiff("n",rs("��¼"),now())
       if dlsj<1 and aqjh_grade<10 then
        kapian="<img src=card/dz04.gif></td><td><font color=red><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�������Ը����ߵ��������˿�,���1���Ӱɣ������˺���!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
end if
      if rs("����")="�ٸ�" then 
        kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�����˿�!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
	  end if
       mtky=mtk(to1)
       if mtky="ok" then   
	      kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"ʹ�����˿�,����һ�ţ������"&to1&"���˳�ȥ!"
    	  call boot(to1,to1&"��"&aqjh_name&"ʹ�������˿�")     
    	else
			kapian=mtky
    	end if
case "���߿�"      
%>
<!--#include file="wnkif.asp"-->
<%
  rs.close
  rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
  if rs("����")="�ٸ�" then
     kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�ö��߿�!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
  end if
   rs.close
   qxky=qxk(to1)
   if qxky="ok" then   
	   conn.execute "update �û� set ��¼=now()+12/(4*144),״̬='��' where ����='" & to1 & "'"
	   kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��"&to1&"ʹ�ö��߿�!ʹ"&to1&"˯����!"
	   call boot(to1,to1&"��"&aqjh_name&"ʹ���˶��߿�")
   else
   		kapian=qxky
   end if
case "��˰��" 
%>
<!--#include file="wnkif.asp"-->
<%
  rs.close   
  rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
  if rs("����")="�ٸ�" then
    kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&",�㲻�ܶԹٸ���Աʹ�ò�˰��!"
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    exit function
  end if
    yl=rs("����")+rs("���")
    if yl<=10000 then
    	kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��"&to1&"ʹ�ò�˰��ʧ��:ԭ��"&to1&"���ϵ�����С��10000��!"
	   	rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	    exit function
	end if
    mhky=mhk(to1)
    if mhky="ok" then   
      yl=int(yl*0.02)
      conn.execute "update �û� set ����=����+"&yl&" where ����='"&aqjh_name&"'"
      if rs("����")>=rs("���") then
        conn.execute "update �û� set ����=����-"&yl&" where ����='"&to1&"'"
      else
        conn.execute "update �û� set ���=���-"&yl&" where ����='"&to1&"'"
      end if
      kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"ʹ�ò�˰��,���"&to1&"��"&yl&"������,ȫ����"&aqjh_name&"����!"
    else
       kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��"&to1&"ʹ�ò�˰��ʧ��!"
       kapian=kapian&mhky
    end if
case "������"
 rs.close
 rs.open "select allvalue,��Ա�ȼ�,�ȼ� from �û� where ����='"&aqjh_name&"'",conn,2,2
 if rs("��Ա�ȼ�")>2 or rs("�ȼ�")>90 then
		Response.Write "<script language=javascript>alert('��ʾ��90�����ϻ�3�����ϻ�Ա����ʹ����������');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
 jhjy=rs("allvalue")
 jhdj=int(sqr(jhjy/50))
 jhadd=((jhdj+1)*(jhdj+1)-jhdj*jhdj)*50
 jhdj=jhdj+1
 jhjy=jhjy+jhadd
 conn.Execute ("update �û� set allvalue="&jhjy&",�ȼ�="&jhdj&" where ����='"&aqjh_name&"'")
 Response.Write "<script language=Javascript>{alert('��["&aqjh_name&"]ʹ��������������������ֵ������"&jhadd&"�㣬ս���ȼ�����1������Ϊ"&jhdj&"��');}</script>"
 kapian=aqjh_name&"ʹ������������"&aqjh_name&"���ݵ�����"&jhadd&"�㣬ս���ȼ�����1��...���Ǻø���ѽ�����ݵ�Ҳ������"
case "������"
 conn.Execute ("update �û� set �书��=�书��+500,������=������+500,������=������+500 where ����='"&aqjh_name&"'")
 Response.Write "<script language=Javascript>{alert('��["&aqjh_name&"]ʹ���˽������������ɹ����书����������������ֵ����500�㣡����');}</script>"
 kapian="����Ƭ��"&aqjh_name&"����ţ���Ҫ��ǿ����Ҫ��ǿ....�ӿڴ����ó�����������"&Application("aqjh_user")&"�ľ��ĵ����£�"&aqjh_name&"����һ�����ѵ�����书����������������ֵ����500�㣡����"
case "���˿�"
	conn.Execute ("update �û� set ɱ����=0  where ����='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('��ʹ���˺��˿���ɱ������0��ɣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##˵�����Ǻ��ˣ��ҽ���û��ɱ��ѽ����</font>"
case "���ۿ�"
	Response.Write "<script>parent.slbox=true;<"&"/script>"	
	Response.Write "<script language=JavaScript>{alert('��ʹ�������ۿ������ڿ��Կ�������˽���ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵�����������ۣ�ʮ���ǧ��ĵط����ܿ����ˡ���</font>"
case "������"
%>
<!--#include file="wnkif.asp"-->
<%
     rs.close
     rs.open "select * from �û� where ����='"&to1&"'",conn,2,2
     jhdj=rs("�ȼ�")
     del1=jhdj*jhdj*50
     del2=(jhdj+1)*(jhdj+1)*50
     delok=del2-del1
     jhjf=rs("allvalue")-delok
 	conn.Execute ("update �û� set �ȼ�=�ȼ�-1,allvalue="&jhjf&" where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�ĵȼ���Ϊ"&jhdj-1&"���ˣ�');}</script>"
	kapian="<img src=card/jiangji.jpg><font color=green><bgsound src=wav/003.mid loop=1><font color=green>����Ƭ��<font color=" & saycolor & ">##��%%ʹ���˽�����,%%�ĵȼ���Ϊ"&jhdj-1&"��������Ҳ��Ϊ����.</font>"
case "�޻���"
rs.close
rs.open "select * from �û� where ����='"&to1&"'",conn,2,2
if rs("ͨ��")="True" then
kapian="<bgsound src=plus/wav/wav/003.mid loop=1><font color=green>����Ƭ��</font><font color=blue>"&aqjh_name&"</font>��<font color=green>"&to1&"</font>ʹ��<font color=red><b>�޻���</b></font>ʧ�ܣ�"&to1&"Ҳ��ͨ�������Ѿ�������޻��ˣ���"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
if rs("����")="�ٸ�" then 
  kapian=""&aqjh_name&"�㲻���ԶԹ���Աʹ�ü޻�������"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
rs.close
rs.open "select ͨ�� from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("ͨ��")="False" then
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>����Ƭ��</font><font color=blue>"&aqjh_name&"</font>��<font color=green>"&to1&"</font>ʹ��<font color=red><b>�޻���</b></font>ʧ�ܣ�"&aqjh_name&"����ͨ��������ô�޻�ѽ����"
        rs.close
set rs=nothing
conn.close
set conn=nothing
exit function
end if
%>
<!--#include file="wnkif.asp"-->
<%
conn.execute "update �û� set ͨ��=True where ����='"&to1&"'"
conn.execute "update �û� set ͨ��=False where ����='"&aqjh_name&"'"
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>����Ƭ��</font>##͵͵ʹ�ÿ�Ƭ������%%��ͨ������ҿ�ɱѽ��"
Case "���ӿ�"
    rs.close
    rs.Open "select hua from �û� where  ����='" & aqjh_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>����Ƭ��</font>##�㲢û���ʻ�,����ʹ�����ӿ�����ȥ����!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>����Ƭ��</font>##����ʻ����������⣬�����¿���!"
                Exit Function
        End If
        '��������
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 5 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update �û� set hua='" & newmyhua & "' where ����='" & aqjh_name & "'"
        kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ�������ӿ�,�õ���5�Ż��֣���ȥ�ֻ��ɣ�</font>"
Case "�ֻ���"
    rs.close
    rs.Open "select hua from �û� where  ����='" & aqjh_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" Or IsNull(myhua) or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>����Ƭ��</font>##�㲢û���ʻ�,����ʹ���ֻ�������ȥ����!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>����Ƭ��</font>##����ʻ����������⣬�����¿���!"
                Exit Function
        End If
        newmyhua = ""
        For I = 0 To 26
                If I >= 2 Then
                temp = Split(zt(I), ";")
                If CLng(temp(0)) > 0 And CLng(temp(0)) < 94 Then
                ss = CLng(temp(0)) + 5 & ";" & temp(1) & ";0"
                newmyhua = newmyhua & ss & "|"
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
        Next
        conn.Execute "update [�û�] set hua='" & newmyhua & "' where ����='" & aqjh_name & "'"
        kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ�����ֻ���,����ˬ�����ۿ��Ż�Ѹ�ٵ����ǣ�</font>"
Case "���̿�"
    rs.close
    rs.Open "select zhu from �û� where  ����='" & aqjh_name & "'", conn,2,2
    myhua = rs("zhu")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>����Ƭ��</font>##�㲢û��������,����ȥ�ؽ�һ��ũ������ʹ�ð�!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>����Ƭ��</font>##����������������⣬��ȥ�ؽ�һ��ũ��!"
                Exit Function
        End If
        '��������
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 2 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update �û� set zhu='" & newmyhua & "' where ����='" & aqjh_name & "'"
        kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##�ͳ������е����̿�,����:�������������ò����׵õ���2ֻ���̣�</font>"
Case "������"
    rs.close
    rs.Open "select zhu from �û� where  ����='" & aqjh_name & "'", conn,2,2
    myhua = rs("zhu")
    If myhua = "" Or IsNull(myhua) or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>����Ƭ��</font>##�㲢û��������,����ȥ�ؽ�һ��ũ������ʹ�ð�!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>����Ƭ��</font>##����������������⣬��ȥ�ؽ�һ��ũ��!"
                Exit Function
        End If
        newmyhua = ""
        For I = 0 To 26
                If I >= 2 Then
                temp = Split(zt(I), ";")
                If CLng(temp(0)) > 0 And CLng(temp(0)) < 94 Then
                ss = CLng(temp(0)) + 5 & ";" & temp(1) & ";0"
                newmyhua = newmyhua & ss & "|"
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
        Next
        conn.Execute "update [�û�] set zhu='" & newmyhua & "' where ����='" & aqjh_name & "'"
        kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ҡ�����е�������,�������������������ɳ��裬�����������������������ţ�</font>"
case else
	Response.Write "<script language=JavaScript>{alert('ϵͳ��û��["&fn1&"]���ֿ�Ƭ,����ʹ�������ˣ�');}</script>"
	Response.End
end select
'ɾ���Լ���Ƭ����¼
conn.execute "update �û� set w5='"&mycard&"' where  ����='"&aqjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function
'��˰��
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"��˰��")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"��˰��",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   'mhk="<br><font color=green>����˰����</font>"&to1&"���ϵ���˰����Ч,��˲���ץ��!"
	   mhk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"��׼��ϲ���̵Ĵ�"&to1&"�Ŀڴ�����Ǯ,���ڴ�ʱ,"&to1&"�ͳ����ϵ���˰��˵,����,����������˰��,�ٺ�!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'���￨
function mzk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"���￨")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���￨",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   'mzk="<br><font color=green>�����￨��</font>"&to1&"���ϵ����￨��Ч,��˲���ץ��!"
	   mzk="<img src=card/myk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&"�ٸ�����׼���ø���������"&to1&"�Ĳ�����,���ڴ�ʱ,"&to1&"�ͳ����ϵ����￨˵,����,�����������￨,�ٺ�!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'���ѿ�
function qxk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"���ѿ�")=0 then
		rs.close
	    qxk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���ѿ�",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   qxk="><img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"�ó�ˮ����˯�ɣ�˯�ɣ��ڴ��ߡ���"&to1&"���Ÿ����۾���ɵ�����Ŀ�����������˵����"&to1&"�ͳ����ϵ����ѿ�,����,�����������ѿ�,�ٺ�!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'���߿�
function mtk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"���߿�")=0 then
		rs.close
	    mtk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���߿�",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   mtk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&aqjh_name&"ʹ���������ţ�׼�����������ж���ȴ��С���ߵ�ʯͷ����"&to1&"��һ�ߺٺٵ�Ц������ѽ����Ҫ���ң�����20�ꡭ��"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'���ܿ�
function wnk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5,ת�� from �û� where ����='"&to1&"'",conn,2,2
	if rs("ת��")>=12 then
		rs.close
		wnk="zstt"
		exit function
	end if
	if iswp(rs("w5"),"���ܿ�")=0 then
		rs.close
	    wnk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���ܿ�",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   wnk="<img src=card/wangneng.jpg></td><td><font color=green style='font-size=9pt'><bgsound src=wav/003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&to1&"��һ�ߺٺٵ�Ц������ѽ�����ܿ������ܿ���һ�����֣��߱����¡���"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>