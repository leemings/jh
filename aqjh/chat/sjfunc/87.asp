<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�����书
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_cs then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����书������Ҫ["&jhdj_cs&"]��,�ſ��Բ�����');}</script>"
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
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�������书��<font color=" & saycolor & ">"+cuan(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�����书
function cuan(fn1,to1)
fn1=abs(fn1)
if fn1<200 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������书һ������200�ģ���̫С����');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ȼ�,grade,�书,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if fn1>50000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����书��Ҫ����50000�ò���');}</script>"
	Response.End
end if
if rs("����")="����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲��ܲ�����');}</script>"
	Response.End
end if
if rs("����")="�ٸ�" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ����˲��ܲ�����');}</script>"
	Response.End
end if
if rs("�书")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������书���㣡');}</script>"
	Response.End
end if
rs.close
rs.open "select �书 FROM �û� WHERE ����='" & to1 &"'",conn
if rs("�书")>1500000 then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ����书�Ѿ�����150���ˣ���������Ҫ�㱣����!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set �书=�书-" & fn1 & " where ����='" & aqjh_name &"'"
conn.execute "update �û� set �书=�书+" & fn1 & " where ����='" & to1 &"'"
cuan= "##����������ͨ�죬����ֱ��,ͷ������ֱð�����ڰ�" & fn1 & "���书������%%�ˣ�%%��ָ�л��"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>	
conn.close
set conn=nothing
end function
%>