<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���⹦�ܡ�wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<9 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ǹ���Ա������ʹ�ñ��⹦�ܣ�');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,��Ա�ȼ� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if sjjh_name<>"һ����" and (rs("����")<10000 or rs("����")<5000000) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����10000������500��ſ���ʹ�ñ��⣡');}</script>"
	Response.End
end if
if rs("��Ա�ȼ�")<2 and sjjh_name<>"һ����" then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ʾ��["&sjjh_name &"]���Ա����2��������ʹ�ñ��⹦�ܣ�');window.close();</script>"
 response.end
end if
conn.execute "update �û� set ����=����-5000000,����=����-10000 where ����='" & sjjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call dianzan(towho)
if len(says)>40 then says=Left(says,40)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=" & saycolor & ">"+mid(says,i+1)+"(<font color=blue><b>"&sjjh_name&"</b></font>)</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)
Application.Lock
Application("sjjh_tltie")=says
Application.UnLock
Response.Write "<script language=JavaScript>{alert('��ʾ��������ɣ�����������10000��������500��');}</script>"
%>