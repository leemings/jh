<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'�������ԡ�wWw.happyjh.com��
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
says="<font color=red>���������ԡ�<font color=" & saycolor & ">"+xsyx()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��������
function xsyx()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����ʱ��,�ݶ�����,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("�ݶ�����")<300 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ķ���������300�����Բ�����');}</script>"
	Response.End
end if
if rs("����")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������100�򲻿��Բ�����');}</script>"
	Response.End
end if
conn.execute "update �û� set �书��=�书��+150,������=������+150,������=������+150,�ݶ�����=�ݶ�����-300,����=����-1000000,����ʱ��=now() where ����='" & sjjh_name &"'"
xsyx="##<bgsound src=wav/dz.wav loop=1>�����������Ե�ѵ��,��������<font color=red>100</font>��,����<font color=red>300</font>��~~�������������书�������<font color=blue><b>150</b></font>�㣬Ŭ���ĽY���������˳�������~~"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>