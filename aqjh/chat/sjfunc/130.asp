<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'������ɽ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>������ħ����<font color=" & saycolor & ">"+xyns()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'������ɽ
function xyns()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select �书,����,�ȼ�,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
xyns=DateDiff("n",rs("����ʱ��"),now())
if xyns<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-xyns
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("�书")<100000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ�书100000���ϲſ��Բ�����');}</script>"
	Response.End
end if
if rs("����") then
	conn.execute "update �û� set ����=����+100,�书=�书-100000,����ʱ��=now() where ����='" & aqjh_name &"'"
	xyns="##<bgsound src=wav/xyns.wav loop=1>ʹ�á�����ħ����������ɽ,ȥ������ɽ�Ժ��书ʧȥ<font color=red>-100000</font>��,��������<font color=red>+100</font>��,##����˵�´��һ�Ҫȥ���Ϸ�!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
