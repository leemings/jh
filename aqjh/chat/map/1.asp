<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<%'�����Ṧ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_im then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ṧ��Ҫ["&jhdj_im&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>�������Ṧ��<font color=" & saycolor & ">"+impart(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�����Ṧ
function impart(fn1,to1)
fn1=int(abs(fn1))
if (fn1<100 or fn1>500000) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ṧ����100�㣬���5��㣡');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �Ṧ FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("�Ṧ")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô����Ṧ�𣿿�����˵��');}</script>"
	Response.End
end if
conn.execute "update �û� set �Ṧ=�Ṧ-"& fn1 &" where ����='" & aqjh_name &"'"
conn.execute "update �û� set �Ṧ=�Ṧ+"& fn1 &" where ����='" & to1 &"'"
impart=aqjh_name & "���Լ���"& fn1 &"���Ṧ���ڸ���<font color=red>" & to1 & "</font>��" & to1 & "�Ṧ���ǣ�"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
