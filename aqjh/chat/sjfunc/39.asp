<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'ŭ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_jhdj<jhdj_nh then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ŭ����Ҫ�ȼ�["&jhdj_nh&"]�ſɲ�������');}</script>"
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
says=nuhou(mid(says,i+1))
towho="���"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ŭ��
function nuhou(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<1000  then
	Response.Write "<script language=JavaScript>{alert('��ʾ������1000������ʹ�ñ���������');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �û� set ����=����-1000 where ����='" & aqjh_name &"'"
nuhou="<marquee height=50 behavior=alternate loop=200 direction=up onmouseover=this.stop(); onmouseout=this.start();><marquee behav onmouseover=this.stop(); onmouseout=this.start();><ior=alternate onmouseover=this.stop(); onmouseout=this.start();>" & fn1 & "(##)" & "</marquee></marquee>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>