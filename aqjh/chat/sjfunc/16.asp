<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��Ǯ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_sq then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ǯ��Ҫ["&jhdj_sq&"]���ſ��Բ�����');}</script>"
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
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>����Ǯ��<font color=" & saycolor & ">"+give(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��Ǯ
function give(fn1,to1)
fn1=int(abs(fn1))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")=0 then
if (fn1<100 or fn1>1000000) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ǯ����100�������100������');}</script>"
	Response.End
end if
if rs("����")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô���Ǯ�𣿿�����˵��');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-" & fn1 & " where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����+" & int(fn1*0.8) & " where ����='" & to1 &"'"
give="##��" & fn1 & "�������͸���%%,���¿ɰ�%%�ֵ�ֱ��,����˵лл,%%���ɸ�������˰%20����<bgsound src=wav/THANKS.wav loop=1>"
else
if fn1<10000 then
give="##:��Ҳ̫С���˰ɣ���Ϊ����ү��������ôһ�㣿"
exit function
elseif fn1>20000000 then
give="##:�������ǲ���ү����Ҳ����һ���־�" & fn1 & "����������2000������˰ɡ��㵱������Ӳ��������ѵĲ����۰���"
exit function
end if
conn.execute "update �û� set ����=����+" &fn1& " where ����='" & to1 &"'"
give="%%�����Ǻ�������̫���ˣ��ж��˲���ү<font color=red>["&aqjh_name&"]</font>���ϵ���"&fn1&"������<img src='img/251.GIF'><bgsound src=wav/THANKS.wav loop=1>"
end if
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','����',now(),'���Ͱ���" & fn1 & "��')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
