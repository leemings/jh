<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('��������Ҫ4���ſ��Բ�����');}</script>"
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
says=ling(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'������
function ling(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('������ʲô���뵷���𣿣�');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ְλ,����,���� FROM �û� WHERE ����='"&aqjh_name&"'",conn,2,2
zhiwei=rs("ְλ")
guojia=rs("����")
if zhiwei<>"����" and zhiwei<>"ة��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�Ǿ�������ة�࣬��������ʹ�ù�͢���');}</script>"
	Response.End
end if
if rs("����")< 100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	ling="[ϵͳ]��[##]˵������һ�ι�������Ҫ����һ���򣡹��������ǹ���������㵹ù����"
	exit function
end if
ling="<table width='75%' border='0' cellspacing='0' cellpadding='0' align='center'><tr><td height='15' bgcolor='#FF3399'><div align='center'><font color='#FFFFFF' size='4'><font color=yellow face='Wingdings'>[</font><font color='yellow'><b>�� �� �� ��</b></font><font color=yellow face='Wingdings'>[</font></font></div></td></tr><tr><td bgcolor='#CCCCCC'><div align='center'><font color='red' size='2'>"&guojia&zhiwei&aqjh_name&"Ի,��������������:<br>"&fn1&"</font></div></td></tr></table>"
conn.execute "update �û� set ����=����-100000 where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>