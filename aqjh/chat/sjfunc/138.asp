<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����ͬip
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<grade_ip then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ["&grade_ip&"]���Ĳſ��Բ鿴ip�ģ�');}</script>"
	Response.End
end if
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
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=sameip(towho)
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function sameip(to1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('��ʾ���������վ��,������������');}</script>"
	Response.End
end if
if to1="����Ե" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����ǽ�������,���ܲ飡����');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'���û�ip
rs.open "select lastip,regip,grade FROM �û� WHERE  ����='" & to1 &"'",conn,2,2
if rs("grade")>aqjh_grade then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��Ϊ���Ǹ߼�����Ա������Ȩ�鿴��');}</script>"
	Response.End
end if
ip1=rs("lastip")   '���ip
ip2=rs("regip")    'ע��ip
'����ͬע��ip
rs.close
rs.open "select * FROM �û� WHERE regip='" &ip2&"'",conn,2,2
if rs.recordcount=0 then
   thename2="û���ҵ��κμ�¼"
else
   thename2="|"
   do while not rs.eof
       name=rs("����")
       thename2=thename2&name&"|"
       rs.movenext
   loop
end if
'����ͬ���ip
rs.close
rs.open "select * FROM �û� WHERE lastip='" &ip1&"'",conn,2,2
if rs.recordcount=0 then
   thename1="û���ҵ��κμ�¼"
else
   thename1="|"
   do while not rs.eof
       name=rs("����")
       thename1=thename1&name&"|"
       rs.movenext
   loop
end if
sameip="[��ip]%%�鵽[<font color=blue><b>"&to1&"</b></font>]��ip��ַ����##��ͬע��IP������¼IP��<br>"
sameip=sameip&"[��ip]ע��IP:"&ip2&"["&thename2&"]<br>[��ip]���IP:"&ip1&"["&thename1&"]"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>n
%>