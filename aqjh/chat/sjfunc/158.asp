<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'�Ȼ��˼�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�����Ȼ��˼䣡');}</script>"
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
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>��ħ��ʦ���Ȼ��˼䡿</font><font color=" & saycolor & ">"+meihuoren(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�Ȼ��˼�
function meihuoren(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ȼ�,����,�Ա�,����,�ݶ�����,ְҵ FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
dj=rs("�ȼ�")
fla=rs("����")
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('���ǳ����˰�!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('����������Ա��!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('�����ж�������Ķ���ÿСʱ��ǰ25���ӣ�����������ʱ�䣡');window.close();}</script>"
	Response.End 
end if
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�Ա�")="��" then
Response.Write "<script language=JavaScript>{alert('�˷���ֻ��Ů����ʹ�ã�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if fla<50000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ��Ҫ50000�����ʩչ�ģ�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<60 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ60��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select �Ա�,���� FROM �û� WHERE ����='" & to1 &"'",conn
xb=rs("�Ա�")
money=int(rs("����")/4)
if xb="Ů" then
Response.Write "<script language=JavaScript>{alert('�˷���ֻ���к������ã�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<1000 then
Response.Write "<script language=JavaScript>{alert('�Է���Ǯ����1000����');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ����,���� FROM �û� WHERE ����='" & to1 &"'",conn
falit=int(rs("����")/4)
if rs("����")="�ٸ�"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ��㲻�ܶԸ߼�����Ա��վ���ط����Ա����!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ�����Գ�������ʲô��������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set ����=����-50000 where  ����='"&aqjh_name&"'"
conn.execute "update �û� set �ݶ�����=�ݶ�����+8,����=����+"&money&" WHERE  ����='"&aqjh_name&"'"
conn.execute "update �û� set ����=����-"&money&" where ����='" & to1 &"'"
meihuoren=aqjh_name & "��" & to1 & "ʩչ<bgsound src=wav/phant030a.wav loop=1>�Ȼ��˼��������" & to1 & "��<img src='img/007.gif'><font color=red>" & aqjh_name & "</font>�����ǵ���ò���Ի����ϵ����Ӳ�֪���������ķ�֮һ,����"&money&"��:)##���ϼ�%%���ϵ�<font color=red>8</font>�����ӣ����˵���ȥ����..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>& "��<img src='img/007.gif'><font color=red>" & aqjh_name & "</font>�����ǵ���ò���Ի����ϵ����Ӳ�֪���������ķ�֮һ,����"&money&"��:)..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>