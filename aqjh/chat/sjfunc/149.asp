<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��Ϊ����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=Session("aqjh_name")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
'������Ϣ
aqjh_roominfo=split(Application("aqjh_room"),";")
nowinroom=session("nowinroom")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
chatroomname=trim(chatinfo(0))
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
'������Ϣ
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
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>����Ϊ������<font color=" & saycolor & ">"+mengzhu()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function mengzhu()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if chatroomname<>"��ս����" then
 mengzhu="[##]��������ֻ����[��ս����]�����ڽ��У�"
 exit function
end if
'�жϷ�������
for i=0 to chatroomnum
 chatroomname=Application("aqjh_chatroomname"&i)
 if chatroomname="��ս����" then
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
	chatinfo=split(aqjh_roominfo(nowinroom),"|")
	useronlinename=Application("aqjh_useronlinename"&nowinroom)
	if onlinenow<>1 then
		Response.Write "<script language=JavaScript>{alert('����:սʤ�����˺���������ɣ�');}</script>"
		Response.End
	end if
 end if
next
'�ж�ʱ��
if Minute(time())<50 then
 		Response.Write "<script language=JavaScript>{alert('����:ʱ�仹û���������������');}</script>"
		Response.End
end if
if Application("aqjh_mengzhu")=aqjh_name then
 		Response.Write "<script language=JavaScript>{alert('����:���Ѿ��������˰�������ʲô����');}</script>"
		Response.End
end if
conn.execute "update �û� set ���=���+200 where ����='"&aqjh_name&"'"
Application("aqjh_mengzhu")=aqjh_name
mengzhu="��ϲ[##]ʤ�ν�������������ְλ���ٸ���������һ����200����ң�"
response.write "<script>alert('��ϲ������������!');</script>"
conn.close
set conn=nothing
end function
%>