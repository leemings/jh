<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'�����Ṧ
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
says="<font color=red>�������Ṧ��<font color=" & saycolor & ">"+lianwu()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�����Ṧ
function lianwu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,�Ṧ,�ȼ�,�Ṧ��,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")<2100 or rs("����")<1100 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����2000������1000�ſ���');}</script>"
	Response.End
end if
if rs("�Ṧ")<rs("�ȼ�")*aqjh_wgsx+3800+rs("�Ṧ��") then
	conn.execute "update �û� set �Ṧ=�Ṧ+50,����=����-2000,����=����-1000,����ʱ��=now() where ����='" & aqjh_name &"'"
	lianwu="##<bgsound src=wav/dz.wav loop=1>��������ϰ��,����ʧȥ-2000,����ʧȥ-1000,�Ṧ����+50,�еñ���ʧ��!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("�Ṧ��")<=rs("�ȼ�")*1400 then
	conn.execute "update �û� set �Ṧ��=�Ṧ��+9,����=����-2000,����=����-1000,����ʱ��=now() where ����='" & aqjh_name &"'"
	lianwu="##<bgsound src=wav/dz.wav loop=1>���������Ṧ,����ʧȥ-2000,����ʧȥ-1000,�Ṧ��������+9,�еñ���ʧ��!"
else
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������ˣ������˼�������');}</script>"
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>