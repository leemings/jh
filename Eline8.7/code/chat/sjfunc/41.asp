<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��ip��wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_grade<grade_ip then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ["&grade_ip&"]���Ĳſ��Բ鿴ip�ģ�');}</script>"
	Response.End
end if
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
says=seeip(towho)
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ip
function seeip(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'���û�ip
rs.open "select lastip,regip FROM �û� WHERE  ����='" & to1 &"'",conn,2,2
ip1=rs("lastip")   '���ip
ip2=rs("regip")    'ע��ip
sip=split(rs("lastip"),".")
sip1=split(rs("regip"),".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
'�����ip�ĵ���
rs.close
rs.open "select c,d FROM z WHERE a<="& num &" and b>="&num,conn,2,2
if rs.eof or rs.bof then
	country="����"
	city="δ֪"
else
	country=rs("c")
	city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
last="����:"& country &" ����:"& city
rs.close
'��ע��ip�ĵ���
num=cint(sip1(0))*256*256*256+cint(sip1(1))*256*256+cint(sip1(2))*256+cint(sip1(3))-1
rs.open "select * FROM z WHERE a<="& num &" and b>="&num,conn,2,2
if rs.eof or rs.bof then
	country="����"
	city="δ֪"
else
	country=rs("c")
	city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
reg="����:"& country &" ����:"& city
seeip="[��ip]����Ա�鵽[<font color=blue><b>"&to1&"</b></font>]������ip��ַΪ:"&ip1&",����Ϊ��"&last&"  ע��ip��ַΪ:"&ip2 &"����Ϊ��"&reg
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>