<%@ LANGUAGE=VBScript codepage ="936" %>

<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'��ȡ�Ṧ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����͵Ǯ��');}</script>"
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
says="<font color=green>����ȡ�Ṧ��<font color=" & saycolor & ">"+demand(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��ȡ�Ṧ
function demand(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,��Ա�ȼ�,�Ṧ,����,grade from �û� where ����='" & aqjh_name &"'" ,conn,2,2
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫʹ����ȡ�Ṧ!');}</script>"
	Response.End
end if
fla=rs("�Ṧ")
if fla<500 then
Response.Write "<script language=JavaScript>{alert('����Ṧ�����޷�ʩչѽ������Ҳ��500�㰡��');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ����ʹ����ȡ!');}</script>"
	Response.End
end if
if rs("�ȼ�")<=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���뽭�����ܶ�����ȡ');}</script>"
	Response.End
end if
if rs("grade")>=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ�Թٸ���ȡ���ҿ���Ҫ��̫���ˣ�����');}</script>"
	Response.End
end if
rs.close
rs.open "select �Ա�,�Ṧ FROM �û� WHERE ����='" & to1 &"'",conn
xb=rs("�Ա�")
yin=int(rs("�Ṧ")/3)
if xb="��" then
Response.Write "<script language=JavaScript>{alert('�˹��ܲ��������е����ϣ�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set �Ṧ=�Ṧ-"&yin&" where ����='" & to1 &"'"
conn.execute "update �û� set �Ṧ=�Ṧ+"&yin&" WHERE  ����='" & aqjh_name &"'"
conn.execute "update �û� set �Ṧ=�Ṧ-500 WHERE  ����='" & aqjh_name &"'"
demand=aqjh_name & "<bgsound src=wav/Phant006.wav loop=1>��ͷ���ԵĶ�" & to1 & "˵����λ�����������Ṧ����̫���˿ɷ��ݽ�һ���Ṧ�����ã�������...˵��<font color=red>" & to1 & "</font>������˼�ش������ó�����֮һ���Ṧ����" & aqjh_name & "˵,�����Ṧ���Ҫ�����õ�ȥ��.���û���. " & aqjh_name & "��˵õ�"&yin&"���Ṧ���Ṧ����500��"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>
