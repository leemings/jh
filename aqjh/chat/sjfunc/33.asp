<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'ת��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_jhdj<jhdj_zz then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ת����Ҫ["&jhdj_zz&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>��ת�ˡ�<font color=" & saycolor & ">"+zzyin(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'ת��
function zzyin(fn1,to1)
fn1=abs(fn1)
if fn1<10000 or fn1>5000000 then
	Response.Write "<script language=JavaScript>{alert('ת������1�����500��');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ��� from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("���")<fn1 then
	Response.Write "<script language=JavaScript>{alert('������ô��Ĵ���𣿣�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �û� set ���=���-"& fn1 &" where ����='" & aqjh_name &"'"
conn.execute "update �û� set ���=���+"& int(fn1*0.95) &" where ����='" & to1 &"'"
zzyin="##���Լ������Ĵ��:"& fn1 &"����ת�˵�%%���������£�%%��ٸ���������5%�˴β����ɹ���"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>