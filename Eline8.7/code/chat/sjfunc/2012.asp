<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'���黷��wWw.happyjh.com��
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
says="<font color=red>�����黷��<font color=" & saycolor & ">"+ptz(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function ptz(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �Ա�,w6,�ȼ�,����,����,ְҵ,���� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
if sjjh_name=to1 or to1="���"  then
 ptz=sjjh_name & "��Ա��˳�����?!�������һ��Լ�ʹ�ö��黷ѽ��"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ʹ�ö��黷������ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"���黷",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������һ�����黷Ҳû��Ү.���������!');</script>"
	response.end
end if
if rs("�ȼ�")<180 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ�ñ����㻹��������Ҫ��180�����ϣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<100000 then
	Response.Write "<script language=JavaScript>{alert('�㷨�������ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �ȼ�,״̬,����,���� FROM �û� WHERE ����='" & to1 &"'",conn
if rs("�ȼ�")>500 then
ptz=to1&" ���һ��������˵����ô�㻹����������𣿣���"
else
rs.close
rs.open "select �ȼ�,w6,����,����,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn
duyao=abate(rs("w6"),"���黷",1)
conn.execute "update �û� set w6='"&duyao&"',����=����-100000 where ����='" & sjjh_name &"'"
conn.execute "update �û� set ����=����/2,����=����/2  where ����='" & to1 &"'"
ptz="<font color=red size=2>" & sjjh_name & "</font>͵͵����" & to1 & " �����˶��黷������" & to1 & "�������ֻ���·��������˶��㡣���Ǿ���!(" & to1 & "�����������½������ɣ�)....</font>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>