<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�Ĳ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�ܹ��Ĳ���');}</script>"
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
says="<font color=green>�����߶Ĳ���<font color=" & saycolor & ">"+dubo(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�Ĳ�
function dubo(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
'���жĳ���Ϣ
if fn1="��Ϣ" then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select top 1 a FROM g WHERE c='��' and b='ׯ��' and f=true",conn,2,2
	if rs.eof or rs.bof then
		yinz="����ׯ�ң���"
	else
		yinz="����ׯ�ң�"&rs("a")
	end if
	rs.close
	rs.open "select top 1 a FROM g WHERE c='��' and b='ׯ��' and f=true",conn,2,2
	if rs.eof or rs.bof then
		dianz="���ׯ�ң���"
	else
		dianz="���ׯ�ң�"&rs("a")
	end if
	rs.close
	rs.open "select top 1 a FROM g WHERE c='��' and b='ׯ��' and f=true",conn,2,2
	if rs.eof or rs.bof then
		dianz="���ׯ�ң���"
	else
		dianz="���ׯ�ң�"&rs("a")
	end if
    rs.close
	rs.open "select top 1 a FROM g WHERE c='��' and b='ׯ��' and f=true",conn,2,2
	if rs.eof or rs.bof then
		dianz="���ׯ�ң���"
	else
		dianz="���ׯ�ң�"&rs("a")
	end if

	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	dars=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='С'")
	xiaors=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	danrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='˫'")
	shuangrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and d<>0")
	durs=tmprs("����")
	dubo="��Ϣ��<img src='../jhimg/sz.gif'>[<font color=blue><b>������</b>,"&yinz&"</font>]Ѻ��<font color=blue>"&dars&"</font>��,ѺС��<font color=blue>"&xiaors&"</font>��,Ѻ����<font color=blue>"&danrs&"</font>�ˣ�Ѻ˫��<font color=blue>"&shuangrs&"</font>��,�ϼƣ�<font color=blue>"&durs&"</font>��,���<font color=blue>"&(10-durs)&"</font>�˿���!"
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	dars=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='С'")
	xiaors=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	danrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='˫'")
	shuangrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and d<>0")
	durs=tmprs("����")
	dubo=dubo&"<img src='../jhimg/sz.gif'>[<font color=blue><b>����</b>,"&dianz&"</font>]Ѻ��<font color=blue>"&dars&"</font>��,ѺС��<font color=blue>"&xiaors&"</font>��,Ѻ����<font color=blue>"&danrs&"</font>�ˣ�Ѻ˫��<font color=blue>"&shuangrs&"</font>��,�ϼƣ�<font color=blue>"&durs&"</font>��,���<font color=blue>"&(10-durs)&"</font>�˿���!"
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	dars=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='С'")
	xiaors=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	danrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='˫'")
	shuangrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and d<>0")
	durs=tmprs("����")
	dubo="��Ϣ��<img src='../jhimg/sz.gif'>[<font color=blue><b>������</b>,"&yinz&"</font>]Ѻ��<font color=blue>"&dars&"</font>��,ѺС��<font color=blue>"&xiaors&"</font>��,Ѻ����<font color=blue>"&danrs&"</font>�ˣ�Ѻ˫��<font color=blue>"&shuangrs&"</font>��,�ϼƣ�<font color=blue>"&durs&"</font>��,���<font color=blue>"&(10-durs)&"</font>�˿���!"
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	dars=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='С'")
	xiaors=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='��'")
	danrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and e='˫'")
	shuangrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true and d<>0")
	durs=tmprs("����")
	dubo="��Ϣ��<img src='../jhimg/sz.gif'>[<font color=blue><b>�Ṧ��</b>,"&yinz&"</font>]Ѻ��<font color=blue>"&dars&"</font>��,ѺС��<font color=blue>"&xiaors&"</font>��,Ѻ����<font color=blue>"&danrs&"</font>�ˣ�Ѻ˫��<font color=blue>"&shuangrs&"</font>��,�ϼƣ�<font color=blue>"&durs&"</font>��,���<font color=blue>"&(10-durs)&"</font>�˿���!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
if aqjh_jhdj<jhdj_db then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ս������["&jhdj_db&"]������������ׯ�ģ�');}</script>"
	Response.End 
end if
if right(fn1,2)<>"��ׯ" or (left(fn1,1)<>"��" and left(fn1,1)<>"��" and left(fn1,1)<>"��" and left(fn1,1)<>"��" and left(fn1,1)<>"��") then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ׯ�������"&fn1&"');}</script>"
	Response.End 
end if
f=Minute(time())
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select ���,����,�Ṧ,allvalue,��Ա�ȼ� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if left(fn1,1)="��" then	
	if rs("���")<20000000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��[������]��Ĵ���2000�򣬲�����ׯ��');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 a,f FROM g WHERE c='��' and b='ׯ��'",conn,2,2
	if rs("f")=true then
		Response.Write "<script language=JavaScript>{alert('��ʾ������["&rs("a")&"]Ϊ[����]ׯ��,�㲻����ׯ����');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "update g set a='"&aqjh_name&"',f=true where c='��' and b='ׯ��' and f=false"
	'conn.execute "insert into g(a,b,c,d,e,f) values ('"&aqjh_name &"','ׯ��','��',0,'ׯ',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<img src='../jhimg/sz.gif'>##�ڴ��л����߶ĳ�[������]������ׯ�ɹ������ڴ�ҿ��Ժ���һ����,��������ѽ����"
end if

if left(fn1,1)="��" then
'r=Day(date())
'if r/5=int(r/5) then
'	rs.close
'	set rs=nothing	
'	conn.close
'	set conn=nothing
'	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ĳ�ֻ������5��10��15��20��25��30�Ž��У�');}</script>"
'	Response.End 
'end if
select case rs("��Ա�ȼ�")
		case 1
			hydian=31250
		case 2
			hydian=90000
		case 3
			hydian=250000
		case 4
			hydian=490000
		case else
			hydian=0
end select	
	if (rs("allvalue")-hydian)<4000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��[����]��Ĵ�㲻��4000�㣬������ׯ��');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 a,f FROM g WHERE c='��' and b='ׯ��'",conn,2,2
	if rs("f")=true then
		Response.Write "<script language=JavaScript>{alert('��ʾ������["&rs("a")&"]Ϊ[����]ׯ��,�㲻����ׯ����');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "update g set a='"&aqjh_name&"',f=true where c='��' and b='ׯ��' and f=false"
'	conn.execute "insert into g(a,b,c,d,e,f) values ('"&aqjh_name &"','ׯ��','��',0,'ׯ',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<img src='../jhimg/sz.gif'>##�ڴ��л����߶ĳ�[����]������ׯ�ɹ������ڴ�ҿ��Ժ���һ����,��������ѽ����"
end if

if left(fn1,1)="��" then	
	if rs("����")<4000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��[������]��ķ�������4000��������ׯ��');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 a,f FROM g WHERE c='��' and b='ׯ��'",conn,2,2
	if rs("f")=true then
		Response.Write "<script language=JavaScript>{alert('��ʾ������["&rs("a")&"]Ϊ[������]ׯ��,�㲻����ׯ����');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "update g set a='"&aqjh_name&"',f=true where c='��' and b='ׯ��' and f=false"
	'conn.execute "insert into g(a,b,c,d,e,f) values ('"&aqjh_name &"','ׯ��','��',0,'ׯ',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<img src='../jhimg/sz.gif'>##�ڴ��л����߶ĳ�[������]������ׯ�ɹ������ڴ�ҿ��Ժ���һ����,��������ѽ����"
end if

if left(fn1,1)="��" then	
	if rs("�Ṧ")<4000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��[�Ṧ��]����Ṧ����4000��������ׯ��');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 a,f FROM g WHERE c='��' and b='ׯ��'",conn,2,2
	if rs("f")=true then
		Response.Write "<script language=JavaScript>{alert('��ʾ������["&rs("a")&"]Ϊ[�Ṧ��]ׯ��,�㲻����ׯ����');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "update g set a='"&aqjh_name&"',f=true where c='��' and b='ׯ��' and f=false"
	'conn.execute "insert into g(a,b,c,d,e,f) values ('"&aqjh_name &"','ׯ��','��',0,'ׯ',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<img src='../jhimg/sz.gif'>##�ڴ��л����߶ĳ�[�Ṧ��]������ׯ�ɹ������ڴ�ҿ��Ժ���һ����,��������ѽ����"
end if

end function
%>

