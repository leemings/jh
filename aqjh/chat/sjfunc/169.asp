<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'��ȸ��
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
says="<font color=red>��ħ��ʦ����ȸ�᡿<font color=" & saycolor & ">"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,�ȼ�,ְҵ,���� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 or fn1>200000 then
Response.Write "<script language=JavaScript>{alert('��ȸ��һ������100��಻�ܳ���200000��');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ʹ�ÿ�ȸ�ᣡ����ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"��ȸ��",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������һ�ÿ�ȸ��Ҳû��Ү.���������!');</script>"
	response.end
end if
if rs("�ȼ�")<jhdj_kql then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ�ñ����㻹��������Ҫ��" & jhdj_kql &"�����ϣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>100  then
	Response.Write "<script language=JavaScript>{alert('��̫̰�ˣ����𳬹�100��');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<30000 then
	Response.Write "<script language=JavaScript>{alert('�㷨�������ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ������,������,�书��,����,w6 FROM �û� WHERE ����='" & aqjh_name &"'" ,conn
if rs("������")>5000000 or rs("������")>5000000 or rs("�书��")>5000000 then
lbsw=aqjh_name & " �������ֵ���ߵ��ˣ���Ҫ���𣿺ٺٺ�.(���������ֵ�еĶ�����5000000��)��"
else
duyao=abate(rs("w6"),"��ȸ��",1)
conn.execute "update �û� set w6='"&duyao&"',������=������+"&fn1&",������=������+"&fn1&",�书��=�书��+"&fn1&",����=����-2000 where ����='" & aqjh_name &"'"
lbsw="����<font color=red size=2>" & aqjh_name & "</font> ʹ���˷����<font color=blue>��ȸ��</font>��ӣ��������书���޸���"&fn1&"��."
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>