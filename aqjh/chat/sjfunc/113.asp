<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'��ɱ
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
says="<font color=red>����ɱ��<font color=" & saycolor & ">"+zs(towho,mid(says,i+1))+"</font>"
towhoway=0
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function zs(to1,fn1)
if aqjh_grade>11 then
   zs=aqjh_name&"��Ϊ�ٸ���Ա����Ȼ�벻������֪��Ϊ�����������ܹ��˵��ٸ��ļ���!"
   exit function
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
sql="update �û� set ����=����+1000,״̬='��' where ����='" & aqjh_name & "'"
conn.execute sql
call boot(aqjh_name,"��ɱ�������ߣ�"&aqjh_name&","&fn1)
conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'" & aqjh_name & "','��ɱ','����')"
zs=aqjh_name&"����������ܲ�����....�񱼵����ֽ�����ߵ������·ܲ�����������ȥ,�Ұ��������˼䱯�磬[��Ϊ�����������أ���������1000�㡣]"
conn.close
set conn=nothing
end function
%>