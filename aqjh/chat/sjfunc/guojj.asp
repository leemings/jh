<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'�鿴��������seegjj
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says="<font color=red>���������⡿<font color=" & saycolor & ">"+seegjj()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�鿴��������seegjj
function seegjj()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ���һ���,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
guojia=rs("����")
if guojia="��" or guojia="δ֪" or guojia="��" or guojia="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㿴�����û�м�����ң����鿴�����أ�');}</script>"
	Response.End
end if
myjj=rs("���һ���")
rs.close
rs.open "select jin FROM guo WHERE g='" & guojia & "'",conn,2,2
bgjj=rs("jin")
seejj="##���ֵĶԱ����Ĺ��ף�<font color=red><b>"&myjj&"</b></font>��,["&guojia&"]�Ĺ���Ϊ��<font color=red><b>"&bgjj&"����</b></font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>