<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���⹦��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<jhdj_ti then
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ҫʹ�ñ��⹦����Ҫ["&jhdj_ti&"]�����ϲſ��ԣ�');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<5000 or rs("����")<5000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����5000������500��ſ���ʹ�ñ��⣡');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-5000000,����=����-5000 where ����='" & aqjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call dianzan(towho)
if len(says)>40 then says=Left(says,40)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<marquee width=100% scrollamount=8>����<font color=" & saycolor & ">"+mid(says,i+1)+"����(<font color=blue><b>"&aqjh_name&"</b></font>)</font></marquee>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)
Application.Lock
Application("aqjh_tltie")=says
Application.UnLock
Response.Write "<script language=JavaScript>{alert('��ʾ��������ɣ�����������5000��������500��');}</script>"
%>