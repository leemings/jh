<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="chatconfig.asp"-->
<%'��Ƭ��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####���䴦��#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(8)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ�ÿ�Ƭ��');}</script>"
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
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=kapian(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��Ƭ
function kapian(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if sjjh_name=to1 and instr(";�����;�ݺ���;��˰��;�Զ���;���ҿ�;������;���˿�;���߿�;���׿�;���˿�;��˰��;����;˥��;����;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�������Լ����в�����');</script>"
	Response.End
end if
if to1="���" and instr("������;���˿�;���߿�;���׿�;���˿�;��˰��;����;˥��;����;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�����ܴ�ҽ��в�����');</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
rs.open "select ��Ա��,w5,���� from �û� where  ����='"&sjjh_name&"'",conn,2,2
mycard=abate(rs("w5"),fn1,1)
select case fn1
case "����"
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�������鸳���ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ����["&fn1&"]���񰡣�������ͬ��...</font>"
case "����"
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�������鸳���ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ����["&fn1&"]���񰡣�������ͬ��...</font>"
case "����"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]����������["&to1&"]�����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]���񰡣�����ͬ��...</font>"
case "˥��"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �û� set sl='˥��',slsj=now()+3 where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]����������["&to1&"]�����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]���񰡣�����ͬ��...</font>"
case "����"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �û� set sl='����',slsj=now()+3 where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]����������["&to1&"]�����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]���񰡣�����ͬ��...</font>"
case "����"
	conn.Execute ("update �û� set sl='��',slsj=now() where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]�ɹ���������鸳��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ����["&fn1&"]���񰡣�88��see tomorrow!</font>"
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
	conn.Execute ("update �û� set sl='"&sl(sss)&"',slsj=now()+3 where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ�ʹ����["&fn1&"]���ڵõ�����["&sl(sss)&"]����');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ����["&fn1&"]���񰡣�������ͬ��...��ô���["&sl(sss)&"]���Լ������ϡ�</font>"
	erase sl
case "�����"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �û� set ����=false,����ʱ��=now() where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('�ɹ������["&to1&"]������״̬��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵ʹ���˽����Ƭ��Ҳ��֪����һλ��ϺҪ��ù��...</font>"
case "�ݺ���"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �û� set ����=int(����/3),����=int(����/3) where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]����������ֻʣԭ����1/3�ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##�����̲����̣��ó��Լ��ݺ�������%%��ͷ�ϴ�ȥ,%%ֻ����ǰһ�ڣ�������������ʧ���..</font>"
case "��Ѫ��"
	conn.Execute ("update �û� set ����=����+50000,����=����+50000  where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]��������������������5��㣡');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ���˲�Ѫ������һ�¿����и��ˣ���������������5��㣡</font>"
case "��Ǯ��"
	conn.Execute ("update �û� set ���=���+88880000  where  ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]�Ĵ��������8888��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ������Ǯ�����Լ���С�ɰ���װ�����ˣ�8888��ѽ��</font>"
case "������"
	conn.Execute ("update �û� set �书=�书+10000  where  ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]ʹ�������������书����1��');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ�������������书���Ǵ�������ǣ�����������Ҫ��̫ƽ�ˣ�</font>"
case "�ӵ㿨"
	conn.Execute ("update �û� set allvalue=allvalue+1000  where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]ʹ���˼ӵ㿨���ݵ�������1000�㣡');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʹ���˼ӵ㿨���ǡ��������ݵ�Ҳ�ӵ㣬�����и�����</font>"
case "�Զ���"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �û� set ����ʱ��=now()  where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]ʹ���˳Զ������������ٱ����ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##ʵ�ڶ�%%����Ϊ������ȥ��ʹ���˳Զ�����%%���һ����������ȥ������û����...</font>"
case "������"
	conn.Execute ("update �û� set ����ʱ��=now()  where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('��["&sjjh_name&"]ʹ���˱����������ڱ����ɹ��ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##:����ţ��ұ����ұ�....�ӿڴ����ó��������������ɹ���</font>"
case "���ҿ�"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �û� set �书=int(�书/3)  where ����='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]ʹ���˵��ҿ������书ֻʣ1/3�ˣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##:����ţ�˭Ҳ��Ҫ�����ң���ҪΪ�������ʹ���˵��ҿ���"&to1&"���书ʧȥ���....</font>"
case "�崿��"
	conn.Execute ("update �û� set ������=0  where ����='" & sjjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('��ʹ�����崿������������0��ɣ�');}</script>"
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##�û����ȥ��ʱ��ѽ����</font>"
case "���Կ�"
	rs.close
	rs.open "select �Ա�,��ż FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
	sex=rs("�Ա�")
	pl=rs("��ż")
	rs.close
    if pl="��" then 
        if sex="��" then
               sql="update �û� set �Ա�='Ů' WHERE ����='" & sjjh_name & "'"
               xb="Ư����Ů��"
         end if 
          if sex="Ů" then 
              sql="update �û� set �Ա�='��' WHERE ����='" & sjjh_name & "'"
              xb="Ӣ�����к�"
          end if
          Set Rs=conn.Execute(sql)  
            bianxi=sjjh_name&"ʹ�ñ��Կ���,������Ը�Գ�,�����"&xb&"!" 
        else
          bianxi="ʹ�ñ��Կ�ʧ��!ԭ��:"&sjjh_name&"���м��ҵ�����,��ô�������ѽ!��������������!Ϊ�˳ͽ���"&sjjh_name&"�����ֵ��°ܻ�����,�ش�û��"&sjjh_name&"�ı��Կ�!"
        end if
	kapian="<table border=0><tr><td><img src=card/bxk.jpg></td><td><font color=green>����Ƭ��<font color=" & saycolor & ">##������ͷ����¡�������Ƚ�ѽ���ұ䣬�ұ䡭��(�����ڱ�ɿ��նۣ�����������������<br>�������"&bianxi&"</font></td></tr></table>"

case "��鿨"
	rs.close
	rs.open "select ��ż FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
	peiou=rs("��ż")
	if peiou="��" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��"&sjjh_name&"������Ǯѽ������û����żѽ��');</script>"
		Response.End
	end if
	conn.Execute ("update �û� set ��ż='��'  where ����='" & sjjh_name &"'")
	rs.close
	rs.open "select ��ż FROM �û� WHERE ����='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update �û� set ��ż='��'  where ����='"&peiou&"'")
	end if
	rs.close
	kapian="<table border=0><tr><td><img src=card/lifen.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">##����ǰ��󣬾����Լ�һ��˼�붷����ʹ������鿨,�����������["&peiou&"]����ˡ���</font></td></tr></table>"
case "���׿�"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&sjjh_name&"�����ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from �û� where ����='"&to1&"'",conn,2,2
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&sjjh_name&"���˼��ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("�Ա�")
    if rs("��ż")<>"��" then
	    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����׿�ʧ��:ԭ����"&to1&"���м��ҵ���</font></td></tr></table>"
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
    rs.open "select * from �û� where ����='"&sjjh_name&"'",conn,2,2
    if rs("��ż")<>"��" then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����׿�ʧ��:ԭ�����Լ��Ѿ����м��ҵ�����!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("�Ա�")=sex then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����׿�ʧ��:ԭ����"&to1&"��"&sjjh_name&"��ͬ��!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update �û� set ��ż='"&sjjh_name&"' where ����='"&to1&"'"
    conn.execute "update �û� set ��ż='"&to1&"' where ����='"&sjjh_name&"'"
    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����׿�,������Ը�Գ�����"&to1&"��Ϊ��!</font></td></tr></table>"
case "���ֿ�"
	rs.close
	rs.open "select ���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
	peiou=rs("����")
	if peiou="��" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��"&sjjh_name&"������Ǯѽ������û������ѽ��');</script>"
		Response.End
	end if
	conn.Execute ("update �û� set ����='��'  where ����='" & sjjh_name &"'")
	rs.close
	rs.open "select ���� FROM �û� WHERE ����='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update �û� set ����='��'  where ����='"&peiou&"'")
	end if
	rs.close
	kapian="<table border=0><tr><td><img src=card/lifen.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">##����ǰ��󣬾����Լ�һ��˼�붷����ʹ���˷��ֿ�,�����������["&peiou&"]�����ˡ���</font></td></tr></table>"
case "���˿�"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&sjjh_name&"�����ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from �û� where ����='"&to1&"'",conn,2,2
    if rs("����")="����" then
		Response.Write "<script language=javascript>alert('��"&sjjh_name&"���˼��ǳ����ˣ�����ˣ���');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("�Ա�")
    if rs("����")<>"��" then
	    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����˿�ʧ��:ԭ����"&to1&"�������˵�����</font></td></tr></table>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("����")="�ٸ�" then 
           kapian="�����ԶԹ���Աʹ�����˿�����"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from �û� where ����='"&sjjh_name&"'",conn,2,2
    if rs("����")<>"��" then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����˿�ʧ��:ԭ�����Լ��Ѿ��������˵�����!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("�Ա�")=sex then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����˿�ʧ��:ԭ����"&to1&"��"&sjjh_name&"��ͬ��!</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update �û� set ����='"&sjjh_name&"' where ����='"&to1&"'"
    conn.execute "update �û� set ����='"&to1&"' where ����='"&sjjh_name&"'"
    kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�����˿�,������Ը�Գ�����"&to1&"��Ϊ����!</font></td></tr></table>"
case "������"
    	wnky=wnk(to1)
	    if wnky<>"ok" then 
			kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
			exit function
	    end if
        rs.close
        rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
        if rs("����")="�ٸ�" then 
            kapian="<table border=0><tr><td><img src=card/xianhao.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&",�㲻�ܶԹٸ���Աʹ�ô�����!</td></tr></table>"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        kapian="<table border=0><tr><td><img src=card/xianhao.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"�ó���������Ĵ�����,��"&to1&"��ץ��!</td></tr></table>"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update �û� set ״̬='��',��¼=now()+3 where ����='" & to1 & "'"
            call boot(to1,to1&"��"&sjjh_name&"ʹ���˴�����")  
        else
           kapian=kapian&mzky
        end if
case "���˿�"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
		exit function
    end if
      rs.close
      rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
      if rs("����")="�ٸ�" then 
        kapian="<table border=0><tr><td><img src=card/dz04.gif></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&",�㲻�ܶԹٸ���Աʹ�����˿�!</td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
	  end if
       mtky=mtk(to1)
       if mtky="ok" then   
	      kapian="<table border=0><tr><td><img src=card/dz04.gif></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"ʹ�����˿�,����һ�ţ������"&to1&"���˳�ȥ!</td></tr></table>"
    	  call boot(to1,to1&"��"&sjjh_name&"ʹ�������˿�")     
    	else
			kapian=mtky
    	end if
case "���߿�"      
  wnky=wnk(to1)
  if wnky<>"ok" then 
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
	exit function
  end if
  rs.close
  rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
  if rs("����")="�ٸ�" then
     kapian="<table border=0><tr><td><img src=card/shuimian.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&",�㲻�ܶԹٸ���Աʹ�ö��߿�!</td></tr></table>"
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
	   kapian="<table border=0><tr><td><img src=card/shuimian.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�ö��߿�!ʹ"&to1&"˯����!</td></tr></table>"
	   call boot(to1,to1&"��"&sjjh_name&"ʹ���˶��߿�")
   else
   		kapian=qxky
   end if
case "��˰��" 
  wnky=wnk(to1)
  if wnky<>"ok" then 
	kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
	exit function
  end if
  rs.close   
  rs.open "SELECT * FROM �û� WHERE  ����='" & to1 & "'",conn,2,2
  if rs("����")="�ٸ�" then
    kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&",�㲻�ܶԹٸ���Աʹ�ò�˰��!</td></tr></table>"
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    exit function
  end if
    yl=rs("����")+rs("���")
    if yl<=10000 then
    	kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�ò�˰��ʧ��:ԭ��"&to1&"���ϵ�����С��10000��!</td></tr></table>"
	   	rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	    exit function
	end if
    mhky=mhk(to1)
    if mhky="ok" then   
      yl=int(yl*0.02)
      conn.execute "update �û� set ����=����+"&yl&" where ����='"&sjjh_name&"'"
      if rs("����")>=rs("���") then
        conn.execute "update �û� set ����=����-"&yl&" where ����='"&to1&"'"
      else
        conn.execute "update �û� set ���=���-"&yl&" where ����='"&to1&"'"
      end if
      kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"ʹ�ò�˰��,���"&to1&"��"&yl&"������,ȫ����"&sjjh_name&"����!</td></tr></table>"
    else
       kapian="<table border=0><tr><td><img src=card/chashui.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��"&to1&"ʹ�ò�˰��ʧ��!</td></tr></table>"
       kapian=kapian&mhky
    end if
case "������"
 rs.close
 rs.open "select allvalue,��Ա�ȼ�,�ȼ� from �û� where ����='"&sjjh_name&"'",conn,2,2
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
 conn.Execute ("update �û� set allvalue="&jhjy&",�ȼ�="&jhdj&" where ����='"&sjjh_name&"'")
 Response.Write "<script language=Javascript>{alert('��["&sjjh_name&"]ʹ��������������������ֵ������"&jhadd&"�㣬ս���ȼ�����1������Ϊ"&jhdj&"��');}</script>"
 kapian=sjjh_name&"ʹ������������"&sjjh_name&"���ݵ�����"&jhadd&"�㣬ս���ȼ�����1��...���Ǻø���ѽ�����ݵ�Ҳ������"
case "����"
 conn.Execute ("update �û� set �书��=�书��+500,������=������+500,������=������+500 where ����='"&sjjh_name&"'")
 Response.Write "<script language=Javascript>{alert('��["&sjjh_name&"]ʹ���˽����������ɹ����书����������������ֵ����500�㣡����');}</script>"
 kapian=sjjh_name&"����ţ���Ҫ��ǿ����Ҫ��ǿ....�ӿڴ����ó���������"&Application("sjjh_user")&"�ľ��ĵ����£�"&sjjh_name&"����һ�����ѵ�����书����������������ֵ����500�㣡����"
case else
	Response.Write "<script language=JavaScript>{alert('ϵͳ��û��["&fn1&"]���ֿ�Ƭ,����ʹ�������ˣ�');}</script>"
	Response.End
end select

'ɾ���Լ���Ƭ����¼
conn.execute "update �û� set w5='"&mycard&"' where  ����='"&sjjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','����','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function


'��˰��
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"��˰��")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"��˰��",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   'mhk="<br><font color=green>����˰����</font>"&to1&"���ϵ���˰����Ч,��˲���ץ��!"
	   mhk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"��׼��ϲ���̵Ĵ�"&to1&"�Ŀڴ�����Ǯ,���ڴ�ʱ,"&to1&"�ͳ����ϵ���˰��˵,����,����������˰��,�ٺ�!</td></tr></table>"
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
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"���￨")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���￨",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   'mzk="<br><font color=green>�����￨��</font>"&to1&"���ϵ����￨��Ч,��˲���ץ��!"
	   mzk="<table border=0><tr><td><img src=card/myk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&"�ٸ�����׼���ø���������"&to1&"�Ĳ�����,���ڴ�ʱ,"&to1&"�ͳ����ϵ����￨˵,����,�����������￨,�ٺ�!</font></td></tr></table>"
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
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"���ѿ�")=0 then
		rs.close
	    qxk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���ѿ�",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   qxk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"�ó�ˮ����˯�ɣ�˯�ɣ��ڴ��ߡ���"&to1&"���Ÿ����۾���ɵ�����Ŀ�����������˵����"&to1&"�ͳ����ϵ����ѿ�,����,�����������ѿ�,�ٺ�!</td></tr></table>"
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
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"���߿�")=0 then
		rs.close
	    mtk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���߿�",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   mtk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&sjjh_name&"ʹ���������ţ�׼�����������ж���ȴ��С���ߵ�ʯͷ����"&to1&"��һ�ߺٺٵ�Ц������ѽ����Ҫ���ң�����20�ꡭ��</td></tr></table>"
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
  conn.open Application("sjjh_usermdb")
  rs.open "select w5 from �û� where ����='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"���ܿ�")=0 then
		rs.close
	    wnk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"���ܿ�",1)
		conn.execute "update �û� set w5='"&tocard&"' where  ����='"&to1&"'"
	   wnk="<table border=0><tr><td><img src=card/mhk.jpg></td><td><font color=green><bgsound src=003.mid loop=1>����Ƭ��<font color=" & saycolor & ">"&to1&"��һ�ߺٺٵ�Ц������ѽ�����ܿ������ܿ���һ�����֣��߱����¡���</td></tr></table>"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>
