<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����������wWw.happyjh.com��
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
says="<font color=red>������������<font color=" & saycolor & ">"+baodou()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��������baodou
function baodou()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ���,grade,�ݶ�����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if trim(rs("���"))="����" and rs("grade")=3 then
	doudi=500
else
	doudi=1000
end if
if rs("�ݶ�����")<doudi then
	Response.Write "<script language=JavaScript>{alert('��ʾ������������������1����������һС��������3����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<=60 then
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ��"& sjjh_name & "����ʱ��δ������ʣ"& ss &"����!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �û� set ����ʱ��=now(),�ݶ�����=�ݶ�����-"& doudi &" where ����='" & sjjh_name &"'"
baodou="##ף���㱩�����������ɹ��������ڿ�ʼ��Ĺ��������ԭ����3����"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>