<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���˷��
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
says="<font color=green>�����˷�ɡ�<font color=" & saycolor & ">"+xingyu(mid(says,i+1))+"</font>"
if instr(says,"��ϲ")=0 then
	towhoway=1
	towho=aqjh_name
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'���˷��
function xingyu(fn1)
if Application("aqjh_xinyu")=0 or isnull(Application("aqjh_xinyu")) then
	randomize()
	rnd1=cint(rnd()*1999)+1
	Application.Lock
	Application("aqjh_xinyu")=rnd1
	Application.UnLock
end if
fn1=int(abs(fn1))
if fn1<0 or fn1>2000 then
	Response.Write "<script language=JavaScript>{alert('�������,���˷��Ϊ��1-2000֮�������');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if fn1=Application("aqjh_xinyu") then
	Application.Lock
	Application("aqjh_xinyu")=0
	Application.UnLock
	conn.execute "update �û� set ����=����+5000000 where ����='" & aqjh_name &"'"
	conn.close
	set rs=nothing
	set conn=nothing
	xingyu="����ַ�ɡ��ϲ##�����˿��ַ�ɸ�����Ʊ�����룺"&fn1&"<img src='img/251.gif'>����500��<img src='img/251.gif'>����ұ�ʾ��ϲ��"
else
	rs.open "select ���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
	if rs("����")<1000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('û��Ǯ�ˣ�û�취�ˣ������ˣ�');}</script>"
		Response.End
	end if
	conn.execute "update �û� set ����=����-2000 where ����='" & aqjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	xingyu="����ַ�ɡ##�����˲�Ʊ��Ϊ���ָ�����ҵ�����ף����룺"&fn1&"û���н�,������һ�Σ����л��ᣡ����500��"
end if
end function
%>