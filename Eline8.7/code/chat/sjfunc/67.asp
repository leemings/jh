<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'������wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
if sjjh_jhdj<15 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����������û����أ�15�������Լ����ã�');}</script>"
	Response.End
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
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green><b>��������</b><font color=" & saycolor & ">"+diuqi(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function diuqi(fn1)
if InStr(Application("sjjh_diuqi"),"|")<>0 then
	temp=split(Application("sjjh_diuqi"),"|")
	if ubound(temp)=1 then
		sj=DateDiff("s",temp(1),now())
		if sj<60 then
			s=60-sj
			Response.Write "<script language=JavaScript>{alert('��ʾ�������˶�������������["&s&"]���ٲ�����');}</script>"
			Response.End
		end if
	else
		sj=DateDiff("s",temp(3),now())
		if sj<60 then
			s=60-sj
			Response.Write "<script language=JavaScript>{alert('��ʾ�������˶�������Ʒ����["&s&"]���ٲ�����');}</script>"
			Response.End
		end if
	end if
end if

if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if isnumeric(fn1) then 
	fn1=int(abs(clng(fn1)))
	if fn1<100000 then 
		Response.Write "<script language=JavaScript>{alert('��ʾ����Ҳ̫С��������Ҫ����10��飡');}</script>"
		Response.End 
	end if
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select ���� from �û� where ����='"&sjjh_name&"'",conn,2,2
	if rs("����")<fn1 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��������ô���Ǯ�𣿣�');}</script>"
		Response.End
	end if
	conn.execute "update �û� set ����=����-"&fn1&" where  ����='"&sjjh_name&"'"
	Application.Lock
		Application("sjjh_diuqi")=fn1&"|"&now()
	Application.UnLock
	diuqi="����ͷǮ���Ƕ�ѽ���ٺ١���##���Լ���ѳ�Ʊ�����������У��ܼ�:<a href='dq.asp' target='d'><b><font color=red>+"&fn1&"</font></b></a>��,������˭�ġ�"
else
	fn1=trim(fn1)
	if left(fn1,1)<>"w" or InStr(fn1,"|")=0 then
		Response.Write "<script language=JavaScript>{alert('��ʾ����ʽ���ԣ����|��Ʒ��');}</script>"
		Response.End 
	end if
	zt=split(fn1,"|")
	if not isnumeric(zt(2)) then 
		Response.Write "<script language=JavaScript>{alert('��ʾ����������������ʹ�����֣�');}</script>"
		Response.End 
	end if
	lb=trim(zt(0))
	wpname=trim(zt(1))
	wusl=abs(int(clng(zt(2))))
	if lb<>"w1" and lb<>"w2" and lb<>"w1" and lb<>"w4" and lb<>"w5" and lb<>"w6" and lb<>"w7" and lb<>"w8" and lb<>"w9" and lb<>"w10" then
		Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ�����ȷ��');}</script>"
		Response.End 
	end if
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select "&lb&" from �û� where ����='"&sjjh_name&"'",conn,3,3
	temp=abate(rs(lb),wpname,wusl)
	conn.execute "update �û� set "&lb&"='"&temp&"' where  ����='"&sjjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
		Application("sjjh_diuqi")=lb&"|"&wpname&"|"&wusl&"|"&now()
	Application.UnLock
	diuqi="##���Լ�����Ʒ[<a href='dq.asp' target='d'><font color=red><b>"&wpname&"</b></font></a>]����"&wusl&"��������·��,˭�쵽��˭�ġ���"
end if
end function
%>