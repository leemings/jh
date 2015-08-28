<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
mycorp=session("yx8_mhjh_usercorp")
lastsavetime=session("yx8_mhjh_userlastsavetime")
paycent=Application("yx8_mhjh_paycent")
nowtime=now()
if not isdate(lastsavetime) then lastsavetime=nowtime
expvalue=clng(datediff("n",lastsavetime,nowtime))*paycent*2
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 等级,积分,会员 from 用户 where 姓名='"&username&"'",conn
mygrade=rst("等级")
jifen=rst("积分")
if mygrade>=100 and mycorp<>"官府" then
conn.Execute "update 用户 set 等级=100 where 姓名='"&username&"'"
elseif mygrade>=120 and mycorp="官府" then
conn.Execute "update 用户 set 等级=120 where 姓名='"&username&"'"
end if
if mygrade=1 then 
grade=1000
elseif mygrade=2 then 
grade=2000
elseif mygrade=3 then 
grade=3000
elseif mygrade=4 then 
grade=4000
elseif mygrade=5 then 
grade=5000
elseif mygrade=6 then 
grade=6000
elseif mygrade=7 then 
grade=7000
elseif mygrade=8 then 
grade=8000
elseif mygrade=9 then 
grade=9000
elseif mygrade=10 then 
grade=10000
elseif mygrade=11 then 
grade=11000
elseif mygrade=12 then 
grade=12000
elseif mygrade=13 then 
grade=13000
elseif mygrade=14 then 
grade=14000
elseif mygrade=15 then 
grade=20000
elseif mygrade=16 then 
grade=25000
elseif mygrade=17 then 
grade=30000
elseif mygrade=18 then 
grade=35000
elseif mygrade=19 then 
grade=40000
elseif mygrade=20 then 
grade=50000
elseif mygrade=21 then 
grade=60000
elseif mygrade=22 then 
grade=70000
elseif mygrade=23 then 
grade=80000
elseif mygrade=24 then 
grade=90000
elseif mygrade=25 then 
grade=100000
elseif mygrade=26 then 
grade=120000
elseif mygrade=27 then 
grade=140000
elseif mygrade=28 then 
grade=160000
elseif mygrade=29 then 
grade=180000
elseif mygrade=30 then 
grade=200000
elseif mygrade=31 then 
grade=250000
elseif mygrade=32 then 
grade=300000
elseif mygrade=33 then 
grade=350000
elseif mygrade=34 then 
grade=400000
elseif mygrade=35 then 
grade=450000
elseif mygrade=36 then 
grade=500000
elseif mygrade=37 then 
grade=550000
elseif mygrade=38 then 
grade=600000
elseif mygrade=39 then 
grade=650000
elseif mygrade=40 then 
grade=700000
elseif mygrade=41 then 
grade=750000
elseif mygrade=42 then 
grade=800000
elseif mygrade=43 then 
grade=850000
elseif mygrade=44 then 
grade=900000
elseif mygrade=45 then 
grade=950000
elseif mygrade=46 then 
grade=1000000
elseif mygrade=47 then 
grade=1050000
elseif mygrade=48 then 
grade=1100000
elseif mygrade=49 then 
grade=1150000
elseif mygrade=50 then 
grade=1200000
elseif mygrade=51 then 
grade=1250000
elseif mygrade=52 then 
grade=1300000
elseif mygrade=53 then 
grade=1350000
elseif mygrade=54 then 
grade=1400000
elseif mygrade=55 then 
grade=1450000
elseif mygrade=56 then 
grade=1550000
elseif mygrade=57 then 
grade=1600000
elseif mygrade=58 then 
grade=1750000
elseif mygrade=59 then 
grade=1800000
elseif mygrade=60 then 
grade=1850000
elseif mygrade=61 then 
grade=1900000
elseif mygrade=62 then 
grade=1950000
elseif mygrade=63 then 
grade=2000000
elseif mygrade=64 then 
grade=2050000
elseif mygrade=65 then 
grade=2100000
elseif mygrade=66 then 
grade=2150000
elseif mygrade=67 then 
grade=2200000
elseif mygrade=68 then 
grade=2250000
elseif mygrade=69 then 
grade=2300000
elseif mygrade=70 then 
grade=2350000
elseif mygrade=71 then 
grade=2400000
elseif mygrade=72 then 
grade=2450000
elseif mygrade=73 then 
grade=2500000
elseif mygrade=74 then 
grade=2550000
elseif mygrade=75 then 
grade=2600000
elseif mygrade=76 then 
grade=2650000
elseif mygrade=77 then 
grade=2700000
elseif mygrade=78 then 
grade=2750000
elseif mygrade=79 then 
grade=2800000
elseif mygrade=80 then 
grade=2850000
elseif mygrade=81 then 
grade=2900000
elseif mygrade=82 then 
grade=2950000
elseif mygrade=83 then 
grade=3000000
elseif mygrade=84 then 
grade=3050000
elseif mygrade=85 then 
grade=3100000
elseif mygrade=86 then 
grade=3150000
elseif mygrade=87 then 
grade=3200000
elseif mygrade=88 then 
grade=3250000
elseif mygrade=89 then 
grade=3300000
elseif mygrade=90 then 
grade=3350000
elseif mygrade=91 then 
grade=3400000
elseif mygrade=92 then 
grade=3450000
elseif mygrade=93 then 
grade=3500000
elseif mygrade=94 then 
grade=3550000
elseif mygrade=95 then 
grade=3600000
elseif mygrade=96 then 
grade=3650000
elseif mygrade=97 then 
grade=3700000
elseif mygrade=98 then 
grade=3800000
elseif mygrade=99 then 
grade=3900000
elseif mygrade=100 then 
grade=4000000
end if
pppp=grade
if jifen>pppp and mygrade<=99 then
conn.Execute "update 用户 set 等级=等级+1 where 姓名='"&username&"'"
kl="<marquee width=100% behavior=alternate scrollamount=8><img src=FLOWER/f40.gif border=0><font color=red>"&username&"</font>经过不段的艰苦修炼,等级得到提升:由<font color=red>"& session("yx8_mhjh_usergrade") &"</font>--><font color=red><b>"& session("yx8_mhjh_usergrade")+1 &"</b></font>级.!</marquee>"
session("yx8_mhjh_usergrade")=mygrade+1
says="<font color=red>【升级消息】</font>"&kl	
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
end if
if rst("会员")=true and yx8_mhjh_fellow=true then expvalue=expvalue*2
rst.Close
set rst=nothing
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select * from 用户 where 姓名='"&username&"'"
Set Rst=conn.Execute(sqlstr)
banli=Rst("配偶")
gj=Rst("攻击")
fy=Rst("防御")
if banli="" then banli="无"
if banli<>"无" then
conn.Execute "update 用户 set 攻击=攻击+"&expvalue*10&",防御=防御+"&expvalue*10&",银两=银两+"&expvalue*200&",精力=精力+"&expvalue*10&",积分=积分+"&expvalue&" where 姓名='"&username&"'"
conn.Execute "update 用户 set 银两=银两+"&expvalue*100&" where 姓名='"&banli&"'"
else
conn.Execute "update 用户 set 攻击=攻击+"&expvalue*10&",防御=防御+"&expvalue*10&",银两=银两+"&expvalue*200&",精力=精力+"&expvalue*10&",积分=积分+"&expvalue&" where 姓名='"&username&"'"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
session("yx8_mhjh_userlastsavetime")=nowtime
chatroomnum=Application("yx8_mhjh_chatroomnum")
maxnosaytime=Application("yx8_mhjh_nosaytime")
onlinelist=Application("yx8_mhjh_onlinelist")
onlinelistubd=ubound(onlinelist)
dim newonlinelist()
dim newonlinename()
newonlinenum=0
for i=1 to chatroomnum
redim preserve newonlinename(i)
newonlinename(i)=";"
next
j=1
for i=1 to onlinelistubd step 6
if datediff("s",onlinelist(i+4),nowtime)<=maxnosaytime then
newonlinenum=newonlinenum+1
newonlinename(onlinelist(i+5))=newonlinename(onlinelist(i+5))&onlinelist(i)&";"
redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5)
newonlinelist(j)=onlinelist(i)
newonlinelist(j+1)=onlinelist(i+1)
newonlinelist(j+2)=onlinelist(i+2)
newonlinelist(j+3)=onlinelist(i+3)
if onlinelist(i)<>username then
newonlinelist(j+4)=onlinelist(i+4)
else
newonlinelist(j+4)=nowtime
end if
newonlinelist(j+5)=onlinelist(i+5)
j=j+6
else
Application.Lock
Application("yx8_mhjh_allonline")=replace(Application("yx8_mhjh_allonline"),";"&onlinelist(i)&";",";")
Application.UnLock
end if
next
Application.Lock
if newonlinenum=0 then
dim index(0)
Application("yx8_mhjh_onlinelist")=index
else
Application("yx8_mhjh_onlinelist")=newonlinelist
end if
Application("yx8_mhjh_allonlinenum")=newonlinenum
for i=1 to chatroomnum
Application("yx8_mhjh_onlinename"&i)=newonlinename(i)
next
Application.UnLock
erase newonlinelist
erase newonlinename
Response.Write "<script language=javascript>setTimeout('this.location.href="&chr(34)&"onlinelist.asp"&chr(34)&"',10000);</script><link rel='stylesheet' href='css2.css'><div align=center><font size='4' color='#CC0000' face='幼圆'><b>"&Application("yx8_mhjh_systemname"&chatroomsn)&"</b></font><br><font color=008800>"&Application("yx8_mhjh_allonlinenum")+3&"</font>人在线<hr></div><body bgcolor='#FFE0F8' background='bg1.gif' oncontextmenu=self.event.returnValue=false><table  width='100%'><tr align=center><td><font color=0000FF size=2>储存完毕会员是普通玩家的二倍</font></td></tr><tr><td><font size=1 color=FF0000><b>当前等级："&mygrade&"</font></td></tr><tr><td align=left><font size=2>经过不懈努力<br>攻击增加<font color=FF0000><b>"&expvalue*10&"</b></font><br>防御增加<font color=FF0000><b>"&expvalue*10&"</b></font><br>积分增加<font color=FF0000><b>"&expvalue&"</b></font><br>精力增加<font color=FF0000><b>"&expvalue*10&"</b></font><br>银两增加<font color=FF0000><b>"&expvalue*200&"</b></font><br>升级还差<font color=FF0000><b>"&pppp-jifen&"</b></font></font>"
if banli<>"无" then Response.Write "<br>伴侣<font color=FF0000>"&banli&"</font>增加银两<font color=FF0000><b>"&expvalue*100&"</b></font></font>"
Response.Write "</font></td></tr></table><body>"
%>