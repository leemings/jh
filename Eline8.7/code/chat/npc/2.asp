<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<%'�Ṧ�ݴ�
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
says=Replace(says,"&amp;","&")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>���Ṧ�ݴ桿<font color=" & saycolor & ">"+keep(mid(says,i+1))+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�Ṧ�ݴ�
function keep(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select �Ṧ��,��������,�Ṧ from �û� where ����='" & sjjh_name &"'",conn,2,2
bankmoney=rs("�Ṧ��")
lastdate=date()-rs("��������")
money=rs("�Ṧ")
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		keep="##�����Ṧ��Ų�û���Ṧ,�Ṧÿ����ϢΪ:0.0001%,��ӭ���!"
	else
		keep="##���Ṧ�����:"& bankmoney &"��,��:"& rs("��������") &"����,��0.0001%��,�Ṧ��������:"& newbankmoney &"���Ṧ!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if money<fn1 then
	Response.Write "<script language=JavaScript>{alert('����������ô����Ṧ������');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
keep="##���Ṧ�����������:"& fn1 &"���Ṧ,�Ṧ�����ִ����Ṧ��:"& newbankmoney+fn1 &"��,�����氡�����հ�!"
conn.execute "update �û� set �Ṧ=�Ṧ-"  & fn1 & ",�Ṧ��="& newbankmoney+fn1 &",��������=date() where ����='" & sjjh_name &"'"

rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
