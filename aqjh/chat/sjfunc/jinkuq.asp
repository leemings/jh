<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'ȡ���
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if Weekday(date())=1 and Hour(time())>=20 and Hour(time())<21 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ϊ����ʱ�䣬����ȡǮ����!');window.close();}</script>"
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
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����ȡ�ҡ�<font color=" & saycolor & ">"+getyin(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ȡ���
function getyin(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ���,��������,��� from �û� where ����='" & aqjh_name &"'",conn,2,2
bankmoney=rs("���")
lastdate=date()-rs("��������")
money=rs("���")
newbankmoney=int(bankmoney+bankmoney*0.00001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getyin="##���ڽ�����Ⲣû�н��,���ÿ����ϢΪ:0.00001%,��ӭ��ȡ!"
	else
		getyin="##�ڽ�����:"& bankmoney &"���,��:"& rs("��������") &"����,��0.00001%��,���������:"& newbankmoney &"���!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ľ��������ô���ң�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getyin="##����ȡ��:"& fn1 &"��,������н��:"& newbankmoney-fn1 &"��,�ú����Ǯ,������!"
conn.execute "update �û� set ���=���+"  & fn1 & ",���="& newbankmoney-fn1 &",��������=date() where ����='" & aqjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>