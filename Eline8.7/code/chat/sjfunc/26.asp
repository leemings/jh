<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��ͽ��wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>����ͽ��<font color=" & saycolor & ">"+stu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��ͽ
function stu(to1)
if trim(Application("sjjh_bais_sf"))<> trim(sjjh_name) then
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1 &"]Ҳû�������Ϊʦ��');}</script>"
	Response.End
end if
if trim(Application("sjjh_bais_td"))<> trim(to1) then
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1 &"]Ҳû�������Ϊʦ��');}</script>"
	Response.End
end if
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
conn.execute "update �û� set ����=����-50000,ʦ��='"& sjjh_name &"',ʦ����Ǯ='0' where ����='"& Application("sjjh_bais_td") &"'"
stu=Application("sjjh_bais_td") & "��##������5����ʦ�ѣ����ǵ�ͷ���ǹ����ģ��������##���Լ�Ϊͽ,ֻҪʦ���ڣ��书����ģ�"
Application.Lock
Application("sjjh_bais_sf")=""
Application("sjjh_bais_td")=""
Application.UnLock
conn.close
set conn=nothing
end function
%>