<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_jj then
	Response.Write "<script language=JavaScript>{alert('����ȯ����Ҫ�����ȼ�["&jhdj_jj&"]���Ĳſ��Բ�����');}</script>"
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
says="<font color=red>�����;��顿<font color=" & saycolor & ">"+jingyan(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function jingyan(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
fn1=abs(fn1)
rs.open "select allvalue,grade FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("allvalue")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û����ô��ĵ�ȯ�޷����ͣ�');}</script>"
	Response.End
end if
if fn1>100 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������涨���ϴ���ȯ���ܳ���100');}</script>"
	Response.End
end if
conn.execute "update �û� set allvalue=allvalue-" & fn1 & " where ����='" & sjjh_name &"'"
conn.execute "update �û� set allvalue=allvalue+" & fn1 & " where ����='" & to1 &"'"
jingyan="##��" & fn1 & "�ľ��鴫����%%,���Լ��Ľ����ȼ������ˣ�����ηѽ��"
'��¼
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','���ھ���','"& jingyan & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>