<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'������wWw.51eline.com��
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
says="<font color=red>��������<font color=" & saycolor & ">"+dazhuo()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function dazhuo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,�ȼ�,������,����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
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
if rs("����")<2100  then
	Response.Write "<script language=JavaScript>{alert('��2000���ϵ������ſ��ԣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<rs("�ȼ�")*sjjh_nlsx+2000+rs("������") then
	conn.execute "update �û� set ����=����+300,����=����-2000,����ʱ��=now() where ����='" & sjjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>������������,����ʧȥ-2000��������+300,�еñ���ʧ��!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("������")<=rs("�ȼ�")*1500 then
	conn.execute "update �û� set ������=������+9,����=����-2000,����ʱ��=now() where ����='" & sjjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>������������,����ʧȥ<font color=red>-2000</font>��,������������<font color=red>+9</font>��,�еñ���ʧ��!"
else
	Response.Write "<script language=JavaScript>{alert('��������������ˣ������˼�������');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>