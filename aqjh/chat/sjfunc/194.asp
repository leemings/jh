<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��ť��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<jhdj_qlcy2s then
	call mess("��ʾ��ʹ�ð�ť����Ҫ["&jhdj_qlcy2&"]���ſ��ԣ�",1)
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=" & saycolor & ">"+titl5(mid(says,i+1))+"</font>"
call chatsay("��ť",towhoway,towho,saycolor,addwordcolor,addsays,says)
function titl5(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����ͷ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
tx=rs("����ͷ��")
if rs("����")<300  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ����Ҫ����300���ú�������ɣ�",1)
end if
conn.execute "update �û� set ����=����-300 where ����='" & aqjh_name &"'"
titl5="<button>����ť�֡�" & fn1 & "  (##)" & "</button>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>