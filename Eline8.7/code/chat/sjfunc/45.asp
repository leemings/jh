<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���ջ����wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####���䴦��#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_jr then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ķ�����ջ�����Ҫ["&grade_jr&"]���ſ��Բ�����');}</script>"
	Response.End
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
says=jrht(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'���ջ���
function jrht(fn1)
fn1=trim(fn1)
if len(fn1)>150 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���������ݲ����Գ���150���ַ� ');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
chatroomname=trim(Application("sjjh_chatroomname"&session("nowinroom")))
sqlstr="SELECT * FROM r where a='"&chatroomname&"'"
rs.Open sqlstr,conn,1,2
rs("k")=fn1&"["&sjjh_name&"]"
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
jrht="<font color=green>�����ջ��⡿<font color=" & saycolor & ">##�ѽ��ջ������ݸ���Ϊ[<font color=blue>"&fn1&"</font>]"
end function
%>
