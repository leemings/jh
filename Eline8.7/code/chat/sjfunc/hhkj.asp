<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'�ػ�ھ���wWw.happyjh.com��
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
says="<font color=red>��һ��ħ����<font color=" & saycolor & ">"+hhkj()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�ػ�ھ�
function hhkj()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select �书��,����,�ȼ�,������,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
hhkj=DateDiff("n",rs("����ʱ��"),now())
if hhkj<(int(rnd*20)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*20)+1)-hhkj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")<50000000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����50000000���ϲſ��Բ�����');}</script>"
	Response.End
end if
if rs("�书��") then
	conn.execute "update �û� set �书��=�书��+500,����=����-50000000,����ʱ��=now() where ����='" & sjjh_name &"'"
	hhkj="##<bgsound src=wav/hhkj.wav loop=1>ʹ�á�һ��ħ�����ػ�ھ�,����ʧȥ<font color=red>-50000000</font>��,�书��������<font color=red>+500</font>��,��˭�������ң�"
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