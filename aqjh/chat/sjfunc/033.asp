<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�ͽ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ�����㲻�ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>��վ���ͽ𿨡�<font color=" & saycolor & ">"+givegold(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�ͽ�
function givegold(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10 or fn1>100) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���ͽ�����10�������100����');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��Ա��,����,�Ṧ,grade FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("��Ա��")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô��Ļ�Ա���𣿿�����˵��');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ͻ�Ա����Ҫ1000�����ӵĳ��طѣ�');}</script>"
	Response.End
end if
if rs("�Ṧ")<50000 then
Response.Write "<script language=JavaScript>{alert('��ʾ�����ͻ�Ա����Ҫ2000���Ṧ�������ѣ���');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set �Ṧ=�Ṧ-2000,��Ա��=��Ա��-" & fn1 & ",����=����-1000 where ����='" & aqjh_name &"'"
conn.execute "update �û� set ��Ա��=��Ա��+" & fn1 & " where ����='" & to1 &"'"
givegold="##��" & fn1 & "����͸���%%[$$redb"&fn1&"$$b]�����˴����Ͳ����ɹ�!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','�ͽ�"&fn1&"��')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
onn.close
set conn=nothing
end function
%>
