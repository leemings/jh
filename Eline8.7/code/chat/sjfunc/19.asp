<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="chatconfig.asp"-->
<%'������wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if fjname="����E��" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���ᱦ�����в�����ʹ��������');}</script>"
	Response.End
end if
if sjjh_jhdj<jhdj_pm then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ["&jhdj_pm&"]���ſ��Բ�����');}</script>"
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
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��������<font color=" & saycolor & ">"+mai(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function mai(fn1,to1)
if to1=sjjh_name or to1=Application("sjjh_automanname") then to1="���"
if InStr(Application("sjjh_mai"),"|")<>0 then
	temp=split(Application("sjjh_mai"),"|")
	if ubound(temp)=7 then
		sj=DateDiff("s",temp(6),now())
		if sj<60 then
			s=60-sj
			Response.Write "<script language=JavaScript>{alert('��ʾ���˴�����û����ɣ���["&s&"]���ٲ�����');}</script>"
			Response.End
		end if	
	end if
end if
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>5 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʽΪ����Ǯ|����|���|��Ʒ��|����|˵�� ');}</script>"
	Response.End 
end if
if not isnumeric(zt(0)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ǯ��ʹ�����֣�');}</script>"
	Response.End 
end if
if zt(1)<>"����" and zt(1)<>"���" then 
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ϊ����������ң��������');}</script>"
	Response.End 
end if
lb=zt(2)
if lb<>"w1" and lb<>"w2" and lb<>"w1" and lb<>"w4" and lb<>"w5" and lb<>"w6" and lb<>"w7" and lb<>"w8" and lb<>"w9" and lb<>"w10" then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ�����ȷ��');}</script>"
	Response.End 
end if
if not isnumeric(zt(4)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ������ʹ�����֣�');}</script>"
	Response.End 
end if
yin=abs(int(clng(zt(0))))
bz=trim(zt(1))
lb=trim(zt(2))
zswupin=trim(zt(3))
wusl=abs(int(clng(zt(4))))
sm=zt(5)
if yin=0 or yin>50000000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������0��5ǧ��֮�䣡');}</script>"
	Response.End 
end if
if wusl=0 or wusl>100 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ����Ӧ����0С��100��');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select "&lb&" from �û� where ����='"&sjjh_name&"'",conn,2,2
if  mywpsl(rs(lb),zswupin)<wusl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&zswupin&"]����������("&wusl&")����');}</script>"
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Application.Lock
''����0����Ǯ1������2�����3����Ʒ��4������5��ʱ��6
Application("sjjh_mai")=sjjh_name&"|"&zt(0)&"|"&zt(1)&"|"&lb&"|"&zt(3)&"|"&zt(4)&"|"&now()&"|"&to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
mai="[##]���Լ�����Ʒ[<b><font color=blue>"&zswupin&"</b></font>]���ó�����������ǮΪ��<font color=blue>"&zt(1)&"</font>[<b><font color=red>"&zt(0)&"</font></b>],"&zt(5)&",����Ʒ������[<font color=blue>"&to1&"</font>]�ġ���<input  type=button value='����' onClick=javascript:;mai"&regjm&".disabled=1;window.open('mai.asp','d') name= mai"&regjm&">"
end function
%>