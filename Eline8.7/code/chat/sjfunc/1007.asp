<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������ť��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_jhdj<jhdj_qlcy2s then
	call mess("��ʾ��ʹ�����°�ť��Ҫ["&jhdj_qlcy2&"]���ſ��ԣ�",1)
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
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=" & saycolor & ">"+titl7(mid(says,i+1))+"</font>"
call chatsay("���°�ť",towhoway,towho,saycolor,addwordcolor,addsays,says)
function titl7(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����ͷ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
tx=rs("����ͷ��")
if rs("����")<400  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ����Ҫ����400���ú�������ɣ�",1)
end if
conn.execute "update �û� set ����=����-400 where ����='" & sjjh_name &"'"
titl7="<center><marquee direction=up behavior=alternate height=100><button>�����°�ť��" & fn1 & "  (##)" & "</button></marquee></center>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>