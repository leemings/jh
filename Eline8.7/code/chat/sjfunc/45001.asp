<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'����������wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����͵Ǯ��');}</script>"
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
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>������������<font color=" & saycolor & ">"+qitaoshu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��������
function qitaoshu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,�ȼ�,����,��Ա�ȼ�,����,����,grade from �û� where ����='" & sjjh_name &"'" ,conn,2,2
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫʹ����������!');}</script>"
	Response.End
end if
fla=rs("����")
if fla<500 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��500�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

rs.close
rs.open "select ����,�ȼ�,����,��Ա�ȼ�,����,grade from �û� where ����='" & to1 &"'" ,conn,2,2
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ����ʹ������!');}</script>"
	Response.End
end if
if rs("�ȼ�")<=15 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���뽭�����ܶ�������');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ�Թ���Ա�������ҿ���Ҫ��̫���ˣ�����');}</script>"
	Response.End
end if
rs.close
rs.open "select �Ա�,���� FROM �û� WHERE ����='" & to1 &"'",conn
xb=rs("�Ա�")
yin=int(rs("����")/10)
if xb="Ů" then
Response.Write "<script language=JavaScript>{alert('��������ֻ��Ů�������ã�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

conn.execute "update �û� set ����=����-"&yin&" where ����='" & to1 &"'"
conn.execute "update �û� set ����=����+"&yin&" WHERE  ����='" & sjjh_name &"'"
conn.execute "update �û� set ����=����-500 WHERE  ����='" & sjjh_name &"'"
qitaoshu=sjjh_name & "<bgsound src=wav/Phant006.wav loop=1>��������ض�" & to1 & "˵����λƯ���ɰ���������õ�СMM<img src='img/79.gif' width='50' height='40'>�ɷ��ݽ�һ��Ǯ�����ã�������...˵��<font color=red>" & to1 & "MM</font>������˼�ش������ó�����֮һ�������ݸ���" & sjjh_name & "˵,�ÿ����ĺ��ӣ���ȥ���Եİɣ����û���. " & sjjh_name & "��˵õ�"&yin&"�����ӣ���������500��"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>
