<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�鿴״̬��wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���鿴״̬��Ҫ����Ա�ſ��Բ�����');}</script>"
	Response.End
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
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>���鿴���ϡ�<font color=" & saycolor & ">����Ա�鿴����:"+getstat(towho)+"</font>"
towhoway=1
towho=sjjh_name
if says<>"" then
	call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
end if

'�鿴״̬
function getstat(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������鿴���˲����ڣ�');}</script>"
	Response.End
end if
getstat=to1&":"& rs("�Ա�") & ",����:" & rs("��ż") & ",����:" & rs("����") & ",����:" &rs("����") & ",���:" &rs("���") &  ",�书:" &rs("�书") & ",����:" & rs("����") & ",����:" & rs("����") & ",����:" & rs("����") & ",����:" & rs("����")&",ip:" & rs("lastip") & ",״̬:" & rs("״̬") & ",����:" & rs("����") & ",����:" & rs("����") & ",�Ĵ���:" & rs("�Ĵ���") & ",Ӯ����:" & rs("Ӯ����") & ",ӮǮ:" & rs("ӮǮ") & ",����:" & rs("����") & ",����ȼ�:" & rs("grade") & ",�ܻ���:" & rs("allvalue") & ",�»���:" & rs("mvalue") & ",��¼����:" & rs("times") & ",��:" & rs("��") & ",ľ:" & rs("ľ") & ",ˮ:" & rs("ˮ") & ",��:" & rs("��") & ",��:" & rs("��") & ",���:" & rs("���") & ",����:" & rs("����") & ",��:" & rs("��Ա��") & ",���:" & rs("���") & ",����:" & rs("����")
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
