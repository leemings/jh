<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�β�yiliao
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
says="<font color=red>�����ֻش���<font color=" & saycolor & ">"+yiliao()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�β�yiliao
function yiliao()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ����,ְҵ,����,�ȼ�,������,������,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
if rs("ְҵ")<>"���" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ְҵ���Ǵ�򣬲��ܲ�����');}</script>"
	Response.End
end if

sj=DateDiff("s",rs("����ʱ��"),now())
if sj<600 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=600-sj
	Response.Write "<script language=JavaScript>{alert('�������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")<150000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����150000�ſ��Բ�����');}</script>"
	Response.End
end if
if rs("����")<rs("�ȼ�")*aqjh_tlsx+5260+rs("������") then
	conn.execute "update �û� set ����=����+3500,����=����+1000,����ʱ��=now() where ����='" & aqjh_name &"'"
	yiliao="��ҽ��٢""##<bgsound src=wav/dz.wav loop=1>���Լ�ʹ����������<font color=red>��������1000</font>��,����Ѹ�ٻָ�<font color=red>+3500</font>��,����ҽ������!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("������")<=rs("�ȼ�")*1800 then
	conn.execute "update �û� set ������=������+5000,����=����+5000,����ʱ��=now() where ����='" & aqjh_name &"'"
	yiliao="��ҽ""##<bgsound src=wav/dz.wav loop=1>���Լ�ʹ����������<font color=red>��������5000</font>��,����Ѹ�ٻָ�<font color=red>+5000</font>��,����ҽ������!!"
else
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������ˣ������˼������ư�');}</script>"
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>