<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'ѧ����ɽ
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
says="<font color=green>��ѧ���гɡ�</font><font color=" & saycolor & ">"+xuechen(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'ѧ����ɽ
function xuechen(fn1)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ʦ��,���� from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("ʦ��")<>fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����˲�����ʦ������û��ʦ����');}</script>"
	Response.End
end if
if rs("����")<2000000 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�ж��������ɽ�ѣ�����˽����ɽ��');}</script>"
	Response.End
end if
conn.execute "update �û� set ʦ��='��' where ����='"&aqjh_name&"'"
xuechen="<font color=red>["&aqjh_name&"]������������["&fn1&"]���Ͻ��£�����ѧ���гɣ������˶�������ɽ�Ѻ���ɽ��...�����ն��Ժ������Ҫ���Լ���!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>