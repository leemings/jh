<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../config.asp"-->
<%'ɨ���ж�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<25 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĵȼ�����25��,Ҳ��ɨ�ƣ�С�İ���Ҳ����!');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��ɨ���ж���</font><font color=" & saycolor & ">"+saohuang(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function saohuang(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
if to1="���" or to1=aqjh_name or to1=Application("aqjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('ɨ�ƶ����д��뿴��ϸ�ˣ���');}</script>"
	Response.End
exit function
end if
if rs("�Ա�")<>"��" then
Response.Write "<script language=JavaScript>{alert('ɨ���ж�����Ů��ִ�У��е���һ�ߣ�');}</script>"
Response.End
end if
%>
<!--#include file="data.asp"-->
<%
sql="select ���� FROM ���� WHERE ������='"&to1&"'"
set rs1=connt.execute(sql)
if rs1.eof or rs1.bof then 
	rs1.close
	set rs1=nothing
	connt.close
	set connt=nothing
Response.Write "<script language=javascript>{alert('��ʾ:����û�д��¹�������Ů��');}</script>" 
Response.End 
end if 
conn.execute "update �û� set ����=����-5000000 where ����='"&to1&"'"
conn.execute "update �û� set ����=����+1000000 where ����='"&aqjh_name&"'"
connt.execute "delete * from ���� where ������='"&to1&"'"
saohuang=aqjh_name&"��һ������Ǵ��Ŵ󵶴�������Ժ��["&to1&"]Χ�����ݺݵ�����һ�٣��ȳ��˱�["&to1&"]������ȥ����Ů��["&to1&"]������500��������["&aqjh_name&"]�õ�100��."
rs.close
set rs=nothing
conn.close
set conn=nothing
	rs1.close
	set rs1=nothing
	connt.close
	set connt=nothing
end function 
%>