<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'�Ȼ��˼��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�����Ȼ��˼䣡');}</script>"
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
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("sjjh_automanname") or towho=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>���Ȼ��˼䡿</font><font color=" & saycolor & ">"+meihuoren(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�Ȼ��˼�
function meihuoren(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �ȼ�,����,�Ա�,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,3,3
dj=rs("�ȼ�")
fla=rs("����")
if rs("����")="�ٸ�"  then
Response.Write "<script language=JavaScript>{alert('���ǹٸ���Ա��!');}</script>"
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
if dj<40 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ40��ս���ȼ�ѽ��');}</script>"
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
conn.execute "update �û� set ����=����-50000 where  ����='"&sjjh_name&"'"
conn.execute "update �û� set ����=����+"&money&" WHERE  ����='"&sjjh_name&"'"
conn.execute "update �û� set ����=����-"&money&" where ����='" & to1 &"'"
meihuoren=sjjh_name & "��" & to1 & "ʩչ<bgsound src=wav/phant030a.wav loop=1>�Ȼ��˼��������" & to1 & "��<img src='img/007.gif'><font color=red>" & sjjh_name & "</font>�����ǵ���ò���Ի����ϵ����Ӳ�֪���������ķ�֮һ,����"&money&"��:)..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>