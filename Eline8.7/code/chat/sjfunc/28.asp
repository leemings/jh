<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'��Ŀ�����wWw.happyjh.com��
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
says="<font color=red>����Ŀ��<font color=" & saycolor & ">"+bimu()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��Ŀ����
function bimu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,�ȼ�,������,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
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
if rs("����")<1100  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����1000�ſ��Բ�����');}</script>"
	Response.End
end if
if rs("����")<rs("�ȼ�")*sjjh_tlsx+5260+rs("������") then
	conn.execute "update �û� set ����=����+3500,����=����-1000,����ʱ��=now() where ����='" & sjjh_name &"'"
	bimu="##<bgsound src=wav/dz.wav loop=1>������������,����ʧȥ<font color=red>-1000</font>��,��������<font color=red>+1500</font>��,�еñ���ʧ��!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("������")<=rs("�ȼ�")*1800 then
	conn.execute "update �û� set ������=������+9,����=����-1000,����ʱ��=now() where ����='" & sjjh_name &"'"
	bimu="##<bgsound src=wav/dz.wav loop=1>������������,����ʧȥ-1000������������+9,�еñ���ʧ��!"
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
