<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'�鿴���Ż����wWw.51eline.com��
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
says="<font color=red>������鿴��<font color=" & saycolor & ">"+seejj()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�鿴���Ż���seejj
function seejj()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ���ɻ���,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
mp=rs("����")
if mp="����" or mp="δ֪" or mp="��" or mp="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻹û�����ɣ����ܲ鿴���ɻ���');}</script>"
	Response.End
end if
myjj=rs("���ɻ���")
rs.close
rs.open "select h FROM p WHERE a='" & mp & "'",conn,2,2
bpjj=rs("h")
seejj="##���ֵĶԱ��ŵĹ��ף�<font color=red><b>"&myjj&"</b></font>��,["&mp&"]�Ļ�����Ϊ��<font color=red><b>"&bpjj&"����</b></font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>