<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'ȡˮ����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if Weekday(date())=1 and Hour(time())>=20 and Hour(time())<21 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ϊ����ʱ�䣬����ȡˮ���Բ���!');window.close();}</script>"
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
says="<font color=green>��ȡˮ���ԡ�<font color=" & saycolor & ">"+getfali(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ȡˮ����
function getfali(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ˮ��,��������,ˮ from �û� where ����='" & aqjh_name &"'",conn,2,2
bankmoney=rs("ˮ��")
lastdate=date()-rs("��������")
money=rs("ˮ")
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getfali="##����ˮ��ƿ��û�д��ˮ��,�ٸ���Ϊ����ˮ��ƿÿ����ϢΪ:0.0001%,��ӭ��ȡ"
	else
		getfali="##��ˮ����:"& bankmoney &"ˮ,��:"& rs("��������") &"����,��0.0001%��,ˮ������:"& newbankmoney &"��ˮ!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������������ô���ˮ����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getfali="##��ˮ��ƿ��ȡ��:"& fn1 &"��ˮ,ˮ��ƿ�ִ��з���:"& newbankmoney-fn1 &"��,�ú�ʹ��,�𱻱�������!"
conn.execute "update �û� set ˮ=ˮ+"  & fn1 & ",ˮ��="& newbankmoney-fn1 &",��������=date() where ����='" & aqjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
