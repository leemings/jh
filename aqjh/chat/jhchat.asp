
<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const2.asp"-->
<!--#include file="../const3.asp"-->
<!--#include file="../const4.asp"-->
<!--#include file="../chk.asp"-->
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
application("npc_sxr")=aqjh_name
application("npc_sxsj")=now()
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(aqjh_name)&" ")=0 then ydl=0
	if ydl=1 and clng(nowinroom)<>i then 
		Session.Abandon
		Response.Redirect "../error.asp?id=140"
		Response.End 
	end if
next 
'房间
chatroomname=trim(Application("aqjh_chatroomname"&session("nowinroom")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_yjdh=1 then Chkyjdh()
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if aqjh_disproxy=1 then  Chkproxy()
Dim SplitReflashPage
Dim DoReflashPage
dim shuaxin_time
DoReflashPage=true
shuaxin_time=chatshuaxin_time
ReflashTime=Now()
if (not isnull(session("ReflashTime"))) and cint(shuaxin_time)>0 and DoReflashPage then
	if DateDiff("s",session("ReflashTime"),Now())<cint(shuaxin_time) then
   	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3><font style='font-size:9pt'>本页面起用了防刷新机制，请不要在<b><font color=ff0000>"&shuaxin_time&"</font></b>秒内连续刷新本页面<BR>正在打开页面，请稍候……</font>"
	response.end
	else
	session("ReflashTime")=Now()
	end if
elseif isnull(session("ReflashTime")) and cint(shuaxin_time)>0 and DoReflashPage then
	Session("ReflashTime")=Now()
end if
'对用户资料处理
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM sm where a='滚动'",conn,2,2
mybanner=rs("c")
jhauto=rs("d")
rs.close
rs.open "select k from r where a='"&chatroomname&"'",conn,3,3
if not(rs.eof and rs.bof) then
	jrht=rs("k")
else
	jrht="欢迎来到["&chatroomname&"]祝你聊的开心~~"
end if
rs.close
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
tjrf=rs("通缉")
newuser=rs("times")-1
if newuser=3 then
newuser=2
end if
aqjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
if tjrf=True then
	jhmp="通缉犯"
else
	jhmp=rs("门派")
end if
nickname=rs("姓名")
jhtx=rs("名单头像")
sex=rs("性别")
jhjh=rs("进化")
hymd=rs("好友名单")
mywife=rs("配偶")
guojia=rs("国家")
zhiwei=rs("职位")
mmp=rs("门派")
myzs=rs("转生")
dj=rs("等级")
hk=rs("花魁")
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
 jhsf="贵宾"
end if
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then 
 jhsf="财神"
end if
if rs("配偶")=Application("aqjh_user") and rs("性别")="女" then
 jhsf="站长夫人"
end if
if Application("aqjh_mengzhu")=aqjh_name then 
 jhsf="武林盟主"
end if
userinfo2=0
mytimes="<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
if aqjh_name=Application("aqjh_user") then
	zhanz="<font color=red>[官府]</font>"
else
	zhanz="<font color=red>[官府]</font>"
end if
if dj<18 then
 aqjh_userinto=aqjh_userinto1
else
 aqjh_userinto=aqjh_userinto2
end if
if hydj>0 then
 aqjh_userinto=aqjh_userinto3
end if
if jhsf="贵宾" then
 aqjh_userinto=aqjh_userinto4
end if
if aqjh_grade=4 then
 aqjh_userinto=aqjh_userinto5
elseif aqjh_grade=5 then
 aqjh_userinto=aqjh_userinto6
end if
if hk<1 or isnull(hk) then
    conn.execute "update 用户 set 花魁=0 where 姓名='"&aqjh_name&"'"
else
    userinfo2=1
    aqjh_userinto=aqjh_userinto7
    conn.execute "update 用户 set 魅力=魅力+18,花魁=花魁-1 where 姓名='"&aqjh_name&"'"
end if
if tjrf=True then
 userinfo2=1
 aqjh_userinto=aqjh_userinto8
end if
if jhsf="财神" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto9
end if
if jhsf="武林盟主" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto10
end if
if aqjh_grade>5 and aqjh_grade<10 then
 aqjh_userinto=aqjh_userinto11
elseif aqjh_grade=10 then
 aqjh_userinto=aqjh_userinto12
end if
if zhiwei="君主" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto13
end if
if zhiwei="丞相" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto14
end if
if zhiwei="将军" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto15
end if
aqjh_userinto=aqjh_userinto&mytimes
if newuser=1 then
	userinfo2=1
	conn.Execute "update 用户 set 银两=银两+5000000,金=金+10,木=木+10,水=水+10,火=火+10,土=土+10,times=times+1 where 姓名='"&aqjh_name&"'"
	Response.Write "<script Language=Javascript>alert('☆★接收新人费★☆\n\n欢迎您："&aqjh_name&"！\n您初次来本江湖，请点确定接收站长赠送的新人费\n银两5百万、金木水火土各10点！\n涨钱卡冰神卡各1张！\n只要您喜欢，这里便是你我们共同的家园了\n在这里，您会不断结识来自四面八方的朋友\n还会有很多新鲜的事物和乐趣等着您！\n祝您在本江湖永远开心、愉快！\n\n开启您的江湖生涯吧 →→ →→');</script>"
      aqjh_userinto="<img src=img/new.gif>##收取了站长送的新人费:银两<font color=#cc0000><b>5百万</b></font>/金木水火土各<font color=#cc0000><b>10点</b></font>，进入聊天室了，大家对新朋友要多多照顾啊!<br>[<font color=ff0000>%%</font>]第一次来到我们江湖请大家多多照顾……"
end if
rs.close
rs.open "select id,a,d,g,h,i,l from vh where b='"&aqjh_name&"' and c=true and j=false",conn
if rs.bof and rs.eof then
 aqjh_userinto=aqjh_userinto
elseif userinfo2=0 then
	aqjh_userinto=rs("d")
	vhwj=rs("h")
	vhname=rs("g")
	id=rs("id")
	vhnj=rs("i")
	sj=DateDiff("d",rs("l"),now())
	randomize timer
	ii=int(rnd()*10)
        if aqjh_grade=10 then
	aqjh_userinto="我们的"&zhanz&"##乘坐"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">，在各路神仙的簇拥下，带着五彩光环来到『"&Application("aqjh_chatroomname")&"』，向大家挥手致意！"&mytimes&"<br><font color=red>【爱情消息】</font>"&zhanz&"%%来到江湖了，请各路英雄欢迎。。。。[%%]魅力大涨1000000点！体力上涨1000000点"
	conn.execute "update 用户 set 魅力=魅力+1000000,体力=体力+1000000 where 姓名='"&aqjh_name&"'"
	end if
        if aqjh_grade=5 then
	aqjh_userinto="##<font color=blue>["&mmp&"掌门]</font>乘坐"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">，在弟子的陪同下，神气活现来到了江湖，大喝一声“小的们，我来了！”"&mytimes&"<br><font color=red>【爱情消息】</font>["&mmp&"]掌门[%%]来到江湖了，请帮中弟子欢迎。。。[%%]魅力大涨10000点！体力上涨10000点"
	conn.execute "update 用户 set 魅力=魅力+10000,体力=体力+10000 where 姓名='"&aqjh_name&"'"
	end if
        if aqjh_grade>5 and aqjh_grade<10 then
	aqjh_userinto="##[<font color=red>官府</font>"&jhsf&"]乘坐"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">，在电闪雷鸣中，杀气腾腾地来到了『"&Application("aqjh_chatroomname")&"』，怒吼道：“洒家来也！”"&mytimes&"<br><font color=red>【爱情消息】</font>[官府"&jhsf&"][%%]来到江湖了，请大家小心从事。。。。[%%]魅力大涨50000点！体力上涨50000点"
	conn.execute "update 用户 set 魅力=魅力+50000,体力=体力+50000 where 姓名='"&aqjh_name&"'"
	end if
     if aqjh_grade>4 then
	sj=30 
	vhnj=50
     else
	if sj>42 or vhnj<1 then
		conn.execute "update vh set j=true where id="&id
                aqjh_userinto="##飘飘然的，来到了『"&Application("aqjh_chatroomname")&"』，“小妹初降仙境，各路神仙有礼！”"
	else
		if sj>35 or vhnj<10 then
			if ii<3 then
				conn.execute "update vh set j=true where id="&id
				aqjh_userinto="##["&mmp&""&jhsf&"]乘坐着自家的"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">要来『"&Application("aqjh_chatroomname")&"』，可是心爱的座驾出了些问题--（肯定是使用过度或年久失修，损坏了）"&mytimes&"<br><font color=red>【爱情消息】</font>[%%]的座驾损坏了，只好跑步进入江湖，满身大汗淋漓。。。。"
			else
				if Isnull(vhname) or vhname="" or vhname="无" then vhname=rs("a")
				aqjh_userinto=replace(aqjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				aqjh_userinto=aqjh_userinto&"<br><font color=red>【爱情消息】</font>众小虾呆呆地看着[%%]，流露出一付羡煞慕煞的表情，几时自己也能有。。。[%%]魅力大涨28点！体力上涨1000点"
				conn.execute "update 用户 set 魅力=魅力+28,体力=体力+1000 where 姓名='"&aqjh_name&"'"
				conn.execute "update vh set i=i-1 where id="&id
			end if
		else
			if ii<2 then
				aqjh_userinto="##["&mmp&""&jhsf&"]乘坐着自家的"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">要来『"&Application("aqjh_chatroomname")&"』，可是碰上交通堵塞，只好一路跑步前进，进入"&Application("aqjh_chatroomname")&"时满身大汗淋漓。"&mytimes&"<br><font color=red>【爱情消息】</font>众小虾看着满身大汗淋漓的[%%]，流露出一付钦佩的表情。。。[%%]魅力上涨18点！"
				conn.execute "update 用户 set 魅力=魅力+18 where 姓名='"&aqjh_name&"'"
			else
				if Isnull(vhname) or vhname="" or vhname="无" then vhname=rs("a")
				aqjh_userinto=replace(aqjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				aqjh_userinto=aqjh_userinto&"<br><font color=red>【爱情消息】</font>众小虾呆呆地看着[%%]，流露出一付羡煞慕煞的表情，几时自己也能有。。。[%%]魅力大涨28点！体力上涨1000点"
				conn.execute "update 用户 set 魅力=魅力+28,体力=体力+1000 where 姓名='"&aqjh_name&"'"
			end if
			conn.execute "update vh set i=i-1 where id="&id
		end if
	end if
    end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
scrname=Request.ServerVariables("SCRIPT_NAME")
if Instr(LCase(scrname),"jhchat.asp")=0 then Response.Redirect "../error.asp?id=002"
Response.Write "<bgsound src=img\luckybug1.mid loop=1 volume=50>"
if Application("aqjh_closedoor")="1" then Response.Redirect "../error.asp?id=100"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("aqjh_disproxy")="1" and (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
t=s & ":" & f & ":" & m
'Session("aqjh_lastsaytime")=sj

if Session("aqjh_inthechat")<>"1" then
if InStr(LCase(Application("aqjh_useronlinename"&nowinroom))," " & LCase(aqjh_name) & " ")<>0 then Response.Redirect "../error.asp?id=300"
Session("aqjh_inthechat")="1"
Session("aqjh_savetime")=now()
Session("aqjh_lasttime")=sj
myzanli=0
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 then myzanli=1
'姓名0,性别1,门派2,身份3,头象4,等级5,id6,会员等级7,暂离8,心情9,配偶10,转生11,国家12,职位13,进化14
myonline = aqjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & aqjh_jhdj& "|" & aqjh_id& "|" & hydj&"|"&myzanli&"|"&"正常"&"|"&mywife&"|"&myzs&"|"&guojia&"|"&zhiwei&"|"&jhjh
Application.Lock
dim newonlinelist()
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")

'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'按门派名字汉语拼音排序
'if yjl=0 and len(onuser(2))<len(jhmp) then			'按门派名字长度，长的在前
'if yjl=0 and len(onuser(2))>len(jhmp) then			'按门派名字长度,短的在前
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'按名字汉语拼音排序
'if yjl=0 and len(onuser(0))<len(aqjh_name) then	'按名字长度，长的在前
'if yjl=0 and len(onuser(0))>len(aqjh_name) then	'按名字长度,短的在前
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if
Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)& " "&aqjh_name & " "
erase  newonlinelist
erase  onlinelist
Application.UnLock
Session("SayCount")=Application("SayCount")
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
		Application("SayCount")=Application("SayCount")+1
		i="SayStr"&YuShu(Application("SayCount"))
		Application(i)=Str
	Application.UnLock()
End Sub
act="进入"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
aqjh_userinto=Replace(aqjh_userinto,"%%","<a href=javascript:parent.sw(\'[" & aqjh_name & "]\'); target=f2>" & aqjh_name &"</a>")
says="<font color=#cc0000>【登陆江湖】</font><font color=008800>" & Replace(aqjh_userinto,"##","<img src="&jhtx&"><a href=javascript:parent.sw(\'[" & aqjh_name & "]\'); target=f2>" & aqjh_name &"</a><font color=#000000>ID:"& aqjh_id &"</b></font>") & "<font color=#000000>[ <font color=blue>性别:</font><font color=red>"& sex &"</font> <font color=blue>等级:</font><font color=red>"& aqjh_jhdj&"</font> <font color=blue>国家:</font><font color=red>"& guojia &"</font> <font color=blue>职位:</font><font color=red>"& zhiwei &"</font> <font color=blue>门派:</font><font color=red>"& jhmp &"</font> <font color=blue>身份:</font><font color=red>"& jhsf &"</font> <font color=blue>配偶:</font><font color=red>"& mywife &"</font> <font color=blue>会员等级:</font><font color=red>"& hydj &"</font> <font color=blue>转生:</font><font color=red>"& myzs &"</font> <font color=blue>进化:</font><font color=red>"& jhjh &"</font>]</font><font class=t>(" & t & ")</font><bgsound src=readonly/okok.wma loop=1>"
if newuser=1 then
randomize()
rnd2=int(rnd*400)+1
says=says & "<input type=button value=领取新人费 onClick=""javascript:xrf"&rnd2&".disabled=1;window.open(\'sjfunc/xrf.asp\')"" name=xrf"&rnd2&" >"
'says=says & "<input type=button value=领取新人费 >"
end if
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
end if
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
hiddenadmin="|"&Application("hidden_admin")&"|"
hiddenadmin=Replace(hiddenadmin,",","|")
if aqjh_grade>=9 then hiddenadmin="||"
randomize()
pic=int(rnd*3)
r=int(rnd*136)
b=int(rnd*136)
pcinfo = Request.ServerVariables("HTTP_USER_AGENT")
If InStr(pcinfo, "NT 5") <> 0 Then
 gdsd = 4
Else
 gdsd = 7
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<style>
BODY{background-image: URL(bbg1.jpg);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}\n");</style>
<script Language="Javascript">
var Banner=new Array();
var wen = new Array();
var huida = new Array();
<%
banners=Split(Trim(mybanner),";")
total=UBound(banners)
for o = 0 to total-1
	temp1=banners(o)
	temp1=replace(temp1,chr(34),"")
	temp1=replace(temp1,chr(13),"")
	 Response.Write "Banner[" & o & "]="& chr(34) & temp1 & chr(34) & ";" & chr(13) & chr(10)
next
erase banners
banners=Split(Trim(jhauto),";")
total=UBound(banners)
for o = 0 to total-1
	temp=Split(Trim(banners(o)),"|")
	temp1=temp(0)
	temp2=temp(1)
	temp1=replace(temp1,chr(34),"")
	temp1=replace(temp1,chr(13),"")
	temp2=replace(temp2,chr(34),"")
	temp2=replace(temp2,chr(13),"")
	 Response.Write "wen[" & o & "]="& chr(34) & temp1 & chr(34) & ";" & chr(13) & chr(10)
 	 Response.Write "huida[" & o & "]="& chr(34) & temp2 & chr(34) & ";" & chr(13) & chr(10)
 	 erase temp
next
erase banners
%>
function shake(n) {if (window.top.moveBy) {for (i = 10; i > 0; i--) {for (j = n; j > 0; j--) {window.top.moveBy(0,i);window.top.moveBy(i,0);window.top.moveBy(0,-i);window.top.moveBy(-i,0);}}}}
if(window!=window.top){window.alert("请使用ie浏览器使用本系统！");top.location.href="../exit.asp"}
if(window.name!="aqjh"){ var i=1;while (i<=50){window.alert("你想作什么呀，黑我？这里是不行的，去别处玩去吧！哈！慢慢点50次！！");i=i+1;}top.location.href="../exit.asp"}
var crm="<%=chatroomname%>",bgc="<%=Application("aqjh_chatcolor")%>",systitle="<%=Application("aqjh_tltie")%>";
var myn="<%=aqjh_name%>",mywife="<%=mywife%>",chatbgcolor="<%=chatbgcolor%>",chatimage="<%=chatimage%>";
var hiddenadmin="<%=hiddenadmin%>",cs=<%=aqjh_grade%>,automan="<%=Application("aqjh_automanname")%>";
var myroom=<%=session("nowinroom")%>,slbox=0,lst=0,tbclu=true,mdcls=true,listfaces=false;
var myguojia="<%=session("guojia")%>",myzhiwei="<%=session("zhiwei")%>";
var showhao=false,myxq=1,mymp=1,jhtx="<%=jhtx%>",showpy="<%=hymd%>",showmp="<%=jhmp%>",headhigh=16;
var showsex="<%=sex%>",showseek="",bc = 1,clsok=0,hang = 0 ;showtype = 0,bgimg=""
var mysayc=0,lastsay=new Array(31);
for(mysayc=0;mysayc<31;mysayc++){lastsay[mysayc]=0;}
var gdsd = <%=gdsd%>;
<%if aqjh_dongtai=1 then%>jsjsstr="<\script src=gg.js></\script>";<%else%>jsjsstr="";<%end if%>
var badword = new Array(<%=aqjh_badword%>);
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/～！·＃￥％…；‘’：“”—＊（　）—＋?－＝、／。，？《》" ;
var writejs=false;
var Maxwrite=100;
var askjs="【系统】已存在脚本引用，这条指令自动清除，请使用清屏指令";
var askqp2="<br><font size=2><font color=red>【提示】</font>对话区超载，10秒钟后自动清屏。</font><br>";
var writeNUM=0;
var askqp="<br><font size=2><font color=red>【提示】</font>请点击<a href='javascript:parent.qp()'>[清]</a>屏节约您的资源，提示三次后自动清屏！</font><br>";
document.write("<title>欢迎来到[<%=Application("aqjh_chatroomname")%>],祝您聊的开心！</title>");
function write(cls){var fsize,lheight;
if(cls==1){fsize=this.f2.document.af.fs.value;lheight=this.f2.document.af.lh.value;}else{fsize='10';lheight='125';}
this.f1.document.open();
this.f1.document.writeln("<html><head><title>对话区</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f1.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"宋体\";CURSOR:url('3.cur');font-size:" + fsize + "pt}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL(imgg/<%=r%>.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f1.document.writeln("<\Script Language=\"JavaScript1.1\">var autoScrollOn=1;var scrollOnFunction;var scrollOffFunction;");
this.f1.document.writeln("function scrollit(){if(!parent.f2.document.af.as.checked){autoScrollOn=0;return true;}else{autoScrollOn=1;StartUp();return true;}}");
this.f1.document.writeln("function scrollWindow(){if(autoScrollOn==1){this.scroll(0,65000);parent.f0.scroll(0,65000);setTimeout('scrollWindow()',200);}}");
this.f1.document.writeln("function scrollOn(){autoScrollOn=1;scrollWindow();}");
this.f1.document.writeln("function scrollOff(){autoScrollOn=0;}");
this.f1.document.writeln("function StartUp(){this.onblur=scrollOnFunction;this.onfocus=scrollOffFunction;scrollWindow();}");
this.f1.document.writeln("scrollOnFunction=new Function('scrollOn()');");
this.f1.document.writeln("scrollOffFunction=new Function('scrollOff()');");
this.f1.document.writeln("StartUp();</\script>");
this.f1.document.writeln("<body background='"+bgimg+"' oncontextmenu=self.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f1.document.writeln("<\script language='javascript' src='QQ.js'><\/script><span class=l><font color=red>【浏览器刷新】</font>热烈欢迎<font color=red>〖"+myn+"〗</font>来到《"+crm+"》！<font class=t>(<%=t%>)</font></span>");
this.f1.document.writeln("<%=roomtemp%><br>");
this.f0.document.open();
this.f0.document.writeln("<html><head><title>分屏显示</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f0.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"宋体\";CURSOR:url('2.cur');font-size:" + fsize + "pt;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL(imgg/<%=b%>.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f0.document.writeln("<body background='"+bgimg+"'oncontextmenu=self.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f0.document.writeln("<\script language='javascript' src='QQ.js'><\/script><span class=l><font color=red>【浏览器刷新】</font>热烈欢迎<font color=red>〖"+myn+"〗</font>来到《"+crm+"》！<font class=t>(<%=t%>)</font></span><br><b>【今日消息】</b><%=jrht%><br>"+jsjsstr);
this.t.location.href="t.asp";parent.DB();parent.mytitle(systitle);}
//sh:s0:名色 s1:色彩 s2:动作 s3:说话者 s4:表情 s5:受话 s6:内容 s7:悄悄话 s8:房间
function sh(s0,s1,s2,s3,s4,s5,s6,s7,s8){
	var show = "",ss="";if(s2=="标题"){parent.mytitle(s6);return;}
	if (myroom != s8){return;}//对房间处理
	if (s3!=myn && s5!=myn && s7 ==1 && slbox==0){return;}//对悄悄话处理
	hang=hang+1;
        if (s3==myn && cs<8 &&  s2!="进入" && s2!="发招"){mysayc=mysayc+1;var testsay=new Date();var thissay=testsay.getTime();
        if(mysayc>30){mysayc=0;}
lastsay[mysayc]=Math.ceil(thissay/1000);
if(mysayc==30){saystime=lastsay[30]-lastsay[0];}else{saystime=lastsay[mysayc]-lastsay[mysayc+1];}
        if(saystime<180){
//   this.f2.document.af.towhoway.checked=false;
   this.d.location.href="bombsay.asp?tow="+s5;
  }} 
	 if(hang>800 && clsok==0){if(confirm("你的屏幕显示的发言行数已经超过了800，考虑到您电脑的性能问题 ，是否清屏？点击“确定”清屏，点击“取消”以后不再出现此提示。")){hang=0;parent.qp();}else{clsok=1;}}
	if (s2=="单挑" && s5==myn && this.f2.document.af.dwtx.checked == true){if(confirm("消息：["+s3+"]向您发出单挑请求，您是否与他单挑！")){this.d.location.href="dantiao.asp?name="+s3+"&yn=1";}else{this.d.location.href="dantiao.asp?name="+s3+"&yn=0";}}//单挑处理
	if(s2=="进入"){
	if (hiddenadmin.indexOf("|"+s3+"|")  !=-1){return;}
	parent.m.location.reload();
	if(s3==mywife && this.f2.document.af.py.checked==true){alert("你的配偶["+mywife+"]上线了………");}
	else{if(showpy.indexOf(s3+"|")!=-1 && this.f2.document.af.py.checked==true && s3!=""){alert("你的好友["+s3+"]上线了………");}}
	}
	if (s2=="警告" && s5==myn){alert("管理员对你进行警告，\n请注意你的言行!");}
	if (s2=="轰炸" && s5==myn){i=1;while(i<1000){alert("你ＴＭＤ小心点，先慢慢点1000下吧！\n下次注意了，再捣乱给你原子弹吃吃!");i=i+1;};top.location.href="../exit.asp"}
	if (s2=="原子弹" && s5==myn){while(true){alert("你ＴＭＤ太不像话了，给我点一辈子吧！\n下次注意了，再捣乱给你核武器吃吃!");};}
	if (s2=="清屏"){if(confirm("消息：管理员为了不影响您聊天，建议您清屏！")){hang=0;parent.qp();}else{return;}}
	if(s2=="踢出"){	parent.m.location.reload();return;}
	if(s2=="退出"){	if (hiddenadmin.indexOf("|"+s3+"|")  !=-1){return;}
	parent.m.location.reload();
	if(s3==this.f2.document.af.towho.value ){//alert("["+s3+"]已经不在聊天室中，说话对象自动设置成“大家”。");
	this.f2.document.af.towho.value="大家";}
	else{if(s3==mywife && this.f2.document.af.py.checked==true){alert("你的配偶["+mywife+"]离开聊天室了………");}
	else{if(showpy.indexOf(s3+"|")!=-1 && this.f2.document.af.py.checked==true && s3!=""){alert("你的好友["+s3+"]离开了聊天室………");}}}
	}
	if(s2=="发招" && (s3==myn || s5==myn) && this.f2.document.af.dwtx.checked==true){parent.shake(1);}//发招震动效果
        if(s2=="轩辕" && (s3==myn || s5==myn) && this.f2.document.af.dwtx.checked==true){parent.shake(2);}//轩辕震动效果
	s1 = s1.substring(0,7);
	if (this.f2.document.af.dwtx.checked != true){s6=s6.replace(".wav","");s6=s6.replace(".mid","")}	//处理不让音乐显示
	show ="<font color=" + s1 + ">"//显示文字的色采
	if(s7==1 && s5!="大家" ){show="【私聊】"+show}
	if(s2!="正常"){show=show+s6+"</font><br>"}//对动作等判断
	else{show = show + "<a href=javascript:parent.sw('["+s3+"]'); target=f2><font color=" + s0 + ">"+s3+"</font></a>"+ s4 + "<a href=javascript:parent.sw('["+s5+"]'); target=f2><font color=" + s0 + ">"+s5+"</font></a>" + "说：" + s6 + "</font><br>";
	if ((s3==myn || s5==myn) && tbclu==false){show="<div  style='background:#66CCCC;'>"+show+"</div>"}//处理全屏的色采
	}
	if (s5==automan){//处理机器人
	dz = "【私聊】"+automan+"对"+s3+"说："+huida[Math.floor(Math.random()*huida.length)]+"<br>";
	for(i=0;i<wen.length;i++){if(s6.search(wen[i]) != -1){dz="【私聊】"+automan+"对"+s3+"说："+huida[i]+"<br>"}}
	show=show+dz;}
	if(tbclu && (myn == s3 || myn == s5)){this.f1.document.writeln(show);}else{this.f0.document.writeln(show);}
writeNUM=writeNUM+1;
if(writeNUM>Maxwrite){writqp();writeNUM=0}}
function writqp()
{
if(Maxwrite>30){
Maxwrite=Maxwrite-30;
writask(askqp);
}
else{
setTimeout('qp();',10000)
writask(askqp2);
Maxwrite=100;}
writeNUM=0
}
function writask(ask2){
if(tbclu==true){this.f1.document.writeln(ask2);}
else{this.f0.document.writeln(ask2);}
}
function qp(){writejs=false;Maxwrite=100;this.f0.location.href="about:blank";this.f1.location.href="about:blank";setTimeout('parent.write(1)',500);}
function scrtx(tx){
jsjsstr="<\script src=data/"+tx+" ></\script>";
if (tx=='0') {listfaces=true;parent.m.location.reload();}
else if (tx=='1') {headhigh=32;listfaces=true;parent.m.location.reload();}
else if (tx=='2') {headhigh=24;listfaces=true;parent.m.location.reload();}
else if (tx=='3') {headhigh=16;listfaces=true;parent.m.location.reload();}
else if (tx=='4') {listfaces=false;parent.m.location.reload();}
else if (tx=='5') {showtype=0;parent.m.location.reload();}
else if (tx=='6') {showtype=1;parent.m.location.reload();}
else if (tx=='7') {showtype=2;parent.m.location.reload();}
else if (tx=='8') {showtype=3;showsex="男";parent.m.location.reload();}
else if (tx=='9') {showtype=3;showsex="女";parent.m.location.reload();}
else if (tx=='10') {myxq=1;parent.m.location.reload();}
else if (tx=='11') {myxq=0;parent.m.location.reload();}
else if (tx=='12') {mymp=1;parent.m.location.reload();}
else if (tx=='13') {mymp=0;parent.m.location.reload();}
else {//this.f1.location.href="about:blank";
setTimeout('parent.write(1)',500);}}
function sw(username){var usna;usna=username.substring(1,username.length-1);this.f2.document.af.towho.value=usna;this.f2.document.af.towho.text=usna;this.f2.document.af.sytemp.focus();return;}
//显示状态
function mdshow(data){
lis = data.split("|");
show="<font color=red>【系统提示】</font><font color=green>["+lis[0] +"]账号:"+lis[6]+"[性别:"+lis[1]+"]等级:"+lis[5];
show=show+"&nbsp;这家伙是NPC电脑玩家，等级太低最好不要惹它！[主人:"+lis[3]+"]</font>"
show=show+"<a href=# onclick=\"window.open('npc_see.asp?id="+ lis[6] +"','seenpc','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=610,height=450');\">查看详细</a><br>";
if(tbclu){
	this.f1.document.writeln(show);
}else{
	this.f0.document.writeln(show);}
}
//显示状态
function wmd(b){this.f3.document.writeln(b);}
function shake(n) {if (window.top.moveBy) {for (i = 10; i > 0; i--) {for (j = n; j > 0; j--) {window.top.moveBy(0,i);window.top.moveBy(i,0);window.top.moveBy(0,-i);window.top.moveBy(-i,0);}}}}
function md1(ren){
//名单自动刷新
        if (this.f2.document.af.mdsx.checked == false){return;}
	this.f3.document.open();
	wmd("<html><head><meta http-equiv='content-type' content='text/html; charset=gb2312'><title>在线用户列表</title><style type='text/css'>");
	wmd("body{CURSOR: url('3.cur');font-family:\"黑体\";font-size:12pt;}td{font-family:\"宋体\";font-size:10pt;line-height:125%;}A{color:#ffffff;text-decoration:none;}A:Hover{color: #FF0000;font-family: \"宋体\"; position: relative; left: 2px; top: 1px; clip:  rect(   )}A:Active {color:#ffffff}.b{color:#ffffff;}.g{color:#00FF00;}.hb{color:#00FFFF;}.hg{color:FF00FF;}.d{font-family:\"宋体\";font-size:10pt;color:#E5E5E5;}.z{font-family:\"宋体\";font-size:10pt;color:orange;}.xinxi{font-family:\"宋体\";font-size:10pt;}.banq{font-family:\"宋体\";font-size:10pt;color:ffff00; filter: DropShadow(Color=000000, OffX=1, OffY=1, Positive=1)}.gf{font-family:\"宋体\";font-size:10pt;color:ff6600;}.gfm{font-size:10pt;color:ff0000;}.zl{color:FFFFFF;text-decoration: line-through;}</style>");
	wmd("<div align=\"center\"><font color=\"#b3d4ff\"<b>『"+crm+"』</b></font><hr size=1 color=b3d4ff>");
	wmd("<font color=\"RED\" style=\"font-size:10.5pt\";font-family:\"黑体\">双击滚屏右键秀</font><br><font color=\"RED\" style=\"font-size:12pt\">--</font> <font color=red><b>"+ren+"</b></font><font color=\"RED\" style=\"font-size:10pt\">人在线</font> <font color=\"RED\" style=\"font-size:12pt\">--</font><br></div>");
	wmd("<div id='Tips' style='position:absolute; left:0; top:0; height: 226px;width:140; display=none;'><IFRAME frameBorder=no height=226px marginHeight=0 marginWidth=0 name=show width=140px scrolling=NO noresize></IFRAME></div>");
	wmd("<div id='myTips' style='position:absolute; left:0; top:0; height: 226px;width:130; display=none;'></div>");
	wmd("<\script language=\"JavaScript\">var NS4=(document.layers);var IE4=(document.all);var win=window;var n=0;function findInPage(str){var txt,i,found;if(str==\"\"){return false;}if(NS4){if(!win.find(str))while(win.find(str,false,true))n++;else{n++;}if(n==0)alert(\"您要的名字没有找到！\");}if(IE4){txt=win.document.body.createTextRange();for(i=0;i<=n && (found=txt.findText(str))!=false;i++){txt.moveStart(\"character\",1);txt.moveEnd(\"textedit\");}if(found){txt.moveStart(\"character\",-1);txt.findText(str);txt.select();txt.scrollIntoView();n++;}else{if(n>0){n=0;findInPage(str);}else{alert(\"您要的名字没有找到！\");}}}return false;}");
	wmd("function s(name){");
	wmd("parent.sw(name);");
	wmd("}");
	wmd("<\/script>");
	wmd("</head><body oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=\"006699\" background=\""+chatimage+"\" bgproperties=\"fixed\">");
	wmd("<\script language=\"JavaScript\">var currentpos,timer;function initialize(){timer=setInterval(\"scrollwindow()\",1);}function sc(){clearInterval(timer);}function scrollwindow(){currentpos=document.body.scrollTop; window.scroll(0,++currentpos);if (currentpos != document.body.scrollTop) sc();}document.onmousedown=sc;document.ondblclick=initialize;function New(para_URL){var URL =new String(para_URL);window.open(URL,'','resizable,scrollbars')}");
	wmd("function ShowTips(strUser,pThis,sex){Hiddenmy();s='<s'+'cript language=javascript src=../jhshow/chatdisp.asp?name='+escape(strUser)+'&id=0&sex='+escape(sex)+'></sc'+'ript>';this.show.document.writeln(s);var pTip = document.all['Tips'].style ;pTip.left = 1 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Showmy(data,pThis){lis = data.split('|');s='<table width=130 border=0 cellpadding=1 cellspacing=0 bgcolor=';if(lis[8]==1){s+='#dedfdf';}else{s+='#ffffe7';}");
	wmd("s+='><tr><td width=44% align=center><img src='+lis[4]+' width=32 height=32></td>';");
	wmd("s+='<td width=56% align=left valign=top><font class=xinxi>ID号:'+lis[6]+'</font><br>';");
	wmd("s+='<div align=left><font class=xinxi>性别:';if(lis[1]=='男'){s+='男';}else{s+='女';}s+='<br>等级:'+lis[5]+'</font></div></td></tr>';");
	wmd("s+='<tr><td colspan=2><font class=xinxi>&nbsp门派:'+lis[2]+'<Br>&nbsp;身份:'+lis[3];if(lis[2]!='NPC'){s+='<BR>&nbsp配偶:'+lis[10]+'<br>&nbsp轮回:'+lis[11]+'次<br>&nbsp国家:'+lis[12]+'<br>&nbsp职位:'+lis[13]+'<br>&nbsp进化:'+lis[14]+'<br>&nbsp;&nbsp;&nbsp;';if(lis[7]==0){s+='非会员';};if(lis[7]==1){s+='<img src=hybz/n.gif>一级暗钻';};if(lis[7]==2){s+='<img src=hybz/b.gif>二级蓝钻';};if(lis[7]==3){s+='<img src=hybz/g.gif>收费黄钻';};if(lis[7]==4){s+='<img src=hybz/r.gif>收费红钻';};if(lis[7]==5){s+='<img src=hybz/rgb.gif>五级特钻';};if(lis[7]==6){s+='<img src=hybz/rgb.gif>六级特钻';};if(lis[7]==7){s+='<img src=hybz/rgb.gif>七级特钻';};if(lis[7]==8){s+='<img src=hybz/rgb.gif>八级特钻';}}s+='</font></td></tr></table>';");




	wmd("myTips.innerHTML = s;var pTip = document.all['myTips'].style ;pTip.left = 6 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Hiddenmy(){var obj = document.all['myTips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function Hidden(){this.show.location.href='about:blank';var obj = document.all['Tips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function getPos(obj,type){var n = 0 ;while(obj!=null){if(type=='top'){n += obj.offsetTop;}else{n += obj.offsetLeft;}obj = obj.offsetParent ;}return n;}");
	wmd("<\/script>");
	wmd("<table width=100% Height=30 border=0 cellspacing=0 cellpadding=0 align='left'>");
	wmd("<form name='search' onSubmit='return findInPage(this.string.value);'><tr><td><br><div align=\"center\"><input name='string' type='text' size=8 onChange='n=0;' style='font-family:宋体;font-size:10pt;background-color:008800;color:FFFFFF;border: 1 double'> <input type='submit' value='查找' style='font-size:10pt;background-color:FF9900;color:FFFFFF;width:30;height:16px;border: 1 double'></form></div></td></tr><tr><td class=banq>");
	wmd("<a href=javascript:parent.sw('[大家]');>大家</a><br>");
        wmd("<a href=javascript:parent.sw('["+automan+"]'); title=\"===========&#13&#10陪你说说话&#13&#10陪你唠唠嗑&#13&#10陪你聊聊天&#13&#10正宗三陪&#13&#10===========\";><font color=FFFFFF>"+automan+"</font></a><br>");}
function md2(data){
//名单自动刷新
        if (this.f2.document.af.mdsx.checked == false){return;}
	var cls,mp,faces,friend,myself,gf;
		lis = data.split("|")
		if(lis[1]=="女"&&lis[7]==0)cls="g";
		if(lis[1]=="男"&&lis[7]==0)cls="b";
		if(lis[1]=="女"&&lis[7]!=0)cls="hg";
		if(lis[1]=="男"&&lis[7]!=0)cls="hb";
	        if(lis[2]=="官府"&&lis[7]!=0)cls="gf";
		if(lis[3]=="掌门")mp="z";else mp="d";
		if(lis[8]==1){cls="zl";mp="d";}
    	if(lis[9]=="正常"){xq="";}else{xq=lis[9]}
	//'0姓名,1性别,2门派,3身份,4头象,5等级,6id,7状态,8暂离开,9心情,10国家,11职位,12进化
	if (showtype==1){showstr=showpy;showseek=lis[0];}else{if (showtype==2){showstr=showmp;showseek=lis[2];}else{if (showtype==3){showstr=showsex;showseek=lis[1];}else{showstr="";showseek="";}}}
	if (showpy.search(lis[0]) != -1){friend="<img src='../jhimg/friend.gif' width='12' height='12'>";}
	else{friend=""}
	if (lis[0]==myn){myself="<img src='../jhimg/self.gif' width='12' height='12'>";}
	else{myself="";}
	if (listfaces==true){faces="<img src='"+lis[4]+"' width='"+headhigh+"' height='"+headhigh+"'>"+friend+myself;}
	else{faces="";}
	if(lis[2]=="官府")gf="<font class=gfm>℡</font>";else gf="";


if(lis[14]=="原始人")jhjh="<font class=z><a target='_blank' title='原始人' href='jhjh/help.htm'><img src='time_star.gif' border='0' ></a></font>";

else if(lis[14]=="岩洞人")jhjh="<font class=z><a target='_blank' title='岩洞人' href='jhjh/help.htm'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="稻草人")jhjh="<font class=z><a target='_blank' title='稻草人' href='jhjh/help.htm'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="木头人")jhjh="<font class=z><a target='_blank' title='木头人' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0' ></a></font>"; 

else if(lis[14]=="小铁人")jhjh="<font class=z><a target='_blank' title='小铁人' href='../twt/jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_star.gif' border='0' ></a></font>";

else if(lis[14]=="白钢人")jhjh="<font class=z><a target='_blank' title='白钢人' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="现代人")jhjh="<font class=z><a target='_blank' title='现代人' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="未来人")jhjh="<font class=z><a target='_blank' title='未来人' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_yueliang.gif' border='0' ></a></font>"; 

else if(lis[14]=="太空人")jhjh="<font class=z><a target='_blank' title='太空人'  href='jhjh/help.htm'><img src='time_yueliang.gif' border='0' ><img src='time_yueliang.gif' border='0' ><img src='time_yueliang.gif' border='0' ></a></font>"; 

else if(lis[14]=="宇宙人")jhjh="<font class=z><a target='_blank' title='宇宙人' href='jhjh/help.htm'><img src='time_sun.gif' border='0' ></a></font>"; 

else if(lis[14]=="幻影人")jhjh="<font class=z><a target='_blank' title='幻影人' href='../twt/jhjh/help.htm'><img src='time_sun.gif' border='0'><img src='time_sun.gif' border='0' ></a></font>";
 
else if(lis[14]=="自由人")jhjh="<font class=z><a target='_blank' title='自由人' href='jhjh/help.htm'><img src='time_sun.gif' border='0'><img src='time_sun.gif' border='0'><img src='time_sun.gif' border='0'></a></font>"; 

else jhjh="";




        if(lis[2]=="NPC"){ss=faces+"<a href=\"JavaScript:parent.mdshow(\'"+data+"\');parent.sw(\'[" + lis[0] + "]\');\"";}
	else{ss=faces+"<a href=\"JavaScript:parent.sw(\'[" + lis[0] + "]\');\"";}
	if(lis[2]!="NPC"){ss=ss+" onmouseover=\"Showmy('" + data + "'," + "this" + ");\" onmouseout=\"Hidden();Hiddenmy();\""}
	if(lis[2]!="NPC"){ss=ss+"oncontextmenu=\"ShowTips('" + lis[0] + "'," + "this" + ",'" + lis[1] + "');\"";}
	else{ss=ss+"oncontextmenu=\"JavaScript:alert('[NPC]不支持江湖秀功能!');\""}
	ss=ss+"><font class=\"" + cls + "\">"+lis[0]+"</font></a>"+gf

             mydj=lis[7]
		if (lis[7]==0){mydj="&nbsp;";}
		if (lis[7]==1){mydj="<img src=hybz/n.gif>";}
		if (lis[7]==2){mydj="<img src=hybz/b.gif>";}
		if (lis[7]==3){mydj="<img src=hybz/g.gif>";}
		if (lis[7]==4){mydj="<img src=hybz/r.gif>";}
		if (lis[7]==5){mydj="<img src=hybz/rgb.gif>";}
	    if (lis[7]==6){mydj="<img src=hybz/rgb.gif>";}
		if (lis[7]==7){mydj="<img src=hybz/rgb.gif>";}
		if (lis[7]==8){mydj="<img src=hybz/rgb.gif>";}
		ss=ss+""+mydj+""


	//是否显示门派
        mypai=lis[2]
        if (lis[2]=="官府"){if (lis[3]=="副掌门"){mypai="官府";}}
        if (lis[2]=="官府"){if (lis[3]=="掌门")mypai="站长";}
        if (lis[3]=="贵宾"){mypai="贵宾";}
        if (lis[3]=="国家"){if (lis[3]=="君主")mypai="君主";}
        if (lis[3]=="财神"){mypai="财神";}
        if (lis[3]=="站长夫人"){mypai="站长夫人";}
        if (lis[3]=="武林盟主"){mypai="武林盟主";}
        if (mymp==1){if(lis[2]=="NPC"){ss=ss+"&nbsp;<font color=FFFFFF>" +mypai+ "</font>";}else{if (lis[3]=="财神" || lis[3]=="武林盟主" || lis[3]=="站长夫人" || lis[3]=="贵宾"|| lis[3]=="君主"){ss=ss+"&nbsp;<font color=red>" +mypai+ "</font>";}else{ss=ss+"&nbsp;<font class=\"" + mp + "\">" +mypai+ "</font>"+jhjh;;}}}
	//是否显示心情
	if (myxq==1){ss=ss+"<font class=xq>&nbsp;"+xq+"</font>";}
	ss=ss+"<br>";
	if (hiddenadmin.indexOf("|"+lis[0]+"|")==-1){
	if (showtype != 0){if (showstr.search(showseek) != -1 || lis[0]==myn){wmd(ss);}}else{wmd(ss);}}}
function md3(){
	if (this.f2.document.af.mdsx.checked == false){return;}
		wmd("<br><div align=\"center\"><a href=http://www.7758530.com/ target=_blank><img src=../logo.gif width=88 height=31 border=0 alt=全力打造精彩江湖与论坛></a><HR size=1 color=b3d4ff><font class=banq style='font-size:9pt'><font color=FFFFFF>男玩家</font> <font color=#00FFFF>男会员</font><br><font color=00FF00>女玩家</font> <font color=FF00FF>女会员</font><br>版本:爱情10.0-完美版<br>程序/美工：永不放弃<br><a href=http://wpa.qq.com/msgrd?V=1&Uin=51726805&Site=永不放弃&Menu=yes target=_blank><img src=QQ.gif width=130 height=29 border=0 alt=有什么建议和意见请与站长联系！><br>");
	if (listfaces==true){wmd("为自己.<img src='../jhimg/self.gif' width='12' height='12'><img src='../jhimg/friend.gif' width='12' height='12'>.为好友</font></div>");}
	wmd("</td></tr></table></body></html>")
	this.f3.document.close();
}
function fc(){rn();setTimeout('parent.write()',0);}
function rn(){this.f2.document.af.username.value=myn;}
function clsay(){if(cs<6){setTimeout("this.f2.document.af.sytemp.value=''",0);}}
function IsBadWord(m){var tmp = "" ;for(var i=0;i<m.length; i++){for(var j=0;j<badstr.length;j++)if(m.charAt(i) == badstr.charAt(j)) break;if(j==badstr.length) tmp += m.charAt(i) ;}for(i=0;i<badword.length;i++) if(tmp.search(badword[i]) != -1) return true;return false;}
function Warning(){this.f2.document.af.sytemp.value='';if(bc > 2) d.location.href="autokick.asp";else{bc++;alert("提示:你说了被系统屏蔽的字眼,三次后自动踢出！");}}
function checksays()
{if(this.f2.document.af.addvalues.checked)
{alert("您目前打开了自动泡点功能，不能发言！")
return false;}
var towho=parent.f2.document.af.towho.value;
var npc=parent.f2.document.af.npc.value;
var sytemp=parent.f2.document.af.sytemp.value;
var isnpc=false;
var sytemp1=sytemp.substring(0,1)
var sytemp2=sytemp.substring(0,4)
if(npc.indexOf(";"+towho+"|")==-1){isnpc=false;}else{isnpc=true;}
if(isnpc==true && sytemp1=="/" && sytemp2!=="/发招$" && sytemp2!=="/召唤$" && sytemp2!=="/加血$"){alert("警告:您不能对npc使用此操作！");this.f2.document.af.towho.value="大家";this.f2.document.af.sytemp.value="";return false;}
if(this.f2.document.af.sytemp.value=="/语音聊天$"){window.open("liao.asp?action="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(IsBadWord(this.f2.document.af.sytemp.value)){Warning();return false;}
var maxlingual=20;
var pos=0;
var lingualnum=0;
var i;
var lingualarr=new Array(maxlingual);
var msg=this.f2.document.af.sytemp.value;
var nowmsg=msg.substr(0,2);
var act;
var sjcz = new Array("/偷钱","/情人","/分手","/求婚","/送豆点","/送金币","/奖励","/哑穴","/点穴","/逮捕","/警告","/坐牢","/监禁","/斩首","/下毒","/吸星大法","/偷钱","/投掷","/发招","/传授内力","/送钱","/册封","/踢人","/传授经验","/转账","/收徒","/炸弹","/原子弹","/赠送","/对赌","/送花","/本派刑法","/赌豆","/双人赌博","/本派罚款","/监禁ip","/封锁ip","/宠物","/宝贝攻击","/如沐春风","/乾坤一掷","/拐骗少女","/拐骗少男","/同归于尽","/屏蔽","/宠物护主","/宝贝护主","/偷取金币","/求签","/对赌银两","/打扑克","/打麻将","/打牌","/出牌","/新人费","/官府奖励","/好友","/执行","/传授体力","/传授武功","/传授道德","/邀请赏花","/轩辕","/小孩","/嫁衣神功","/召唤术","/治愈术","/抢劫令"); 
if(lingualnum<maxlingual){lingualarr[lingualnum]=this.f2.document.af.sytemp.value;lingualnum++;}
else
{for (i=0;i<maxlingual;i++){lingualarr[i]=lingualarr[i+1];}lingualarr[i]=this.f2.document.af.sytemp.value;}
var pos=lingualnum;
var cmd=msg.substring(0,msg.length>1?1:msg.length);
var cmd1=msg.substring(0,msg.length>2?2:msg.length);
var saystemp;
saystemp=this.f2.document.af.sytemp.value
//while(saystemp.indexOf(" ") != -1 ){
//  saystemp=saystemp.replace(" ","")
//  }
var aftowho = this.f2.document.af.towho.value
if(IsBadWord(saystemp)){Warning();return false;}
for(i=0;i<sjcz.length;i++)
if(saystemp.search(sjcz[i]) != -1){ 
if(aftowho== "大家" ||aftowho == myn || aftowho==automan ) {
alert("提示：操作对象？"+sjcz[i]+aftowho+automan);
this.f2.document.af.sytemp.value="";
return false;
}}
if(cmd=="/"& cmd1!="//")
{var spacepoint=msg.indexOf("$");
var spacepoint=(spacepoint==-1?msg.length:spacepoint);
var cmd=msg.substring(1,spacepoint);
switch(cmd){
case "主题":act="sjfunc/1.asp";break;case "点穴":act="sjfunc/2.asp";break; 
case "解穴":act="sjfunc/3.asp";break;case "逮捕":act="sjfunc/4.asp";break; 
case "坐牢":act="sjfunc/5.asp";break;case "监禁":act="sjfunc/6.asp";break; 
case "释放":act="sjfunc/7.asp";break;case "斩首":act="sjfunc/8.asp";break; 
case "警告":act="sjfunc/9.asp";break;case "下毒":act="sjfunc/10.asp";break; 
case "偷钱":act="sjfunc/11.asp";break;case "吸星大法":act="sjfunc/12.asp";break;  
case "投掷":act="sjfunc/13.asp";break;case "发招":act="sjfunc/14.asp";break; 
case "传授内力":act="sjfunc/15.asp";break;case "送钱":act="sjfunc/16.asp";break; 
case "赠送":act="sjfunc/17.asp";break;case "状态":act="sjfunc/18.asp";break; 
case "拍卖":act="sjfunc/19.asp";break;case "册封":act="sjfunc/20.asp";break; 
case "公告":act="sjfunc/21.asp";break;case "踢人":act="sjfunc/22.asp";break; 
case "心动":act="sjfunc/23.asp";break;case "放大":act="sjfunc/24.asp";break; 
case "拜师":act="sjfunc/25.asp";break;case "收徒":act="sjfunc/26.asp";break; 
case "打坐":act="sjfunc/27.asp";break;case "闭目":act="sjfunc/28.asp";break; 
case "传授经验":act="sjfunc/29.asp";break;case "练武":act="sjfunc/30.asp";break; 
case "存钱":act="sjfunc/31.asp";break;case "取钱":act="sjfunc/32.asp";break; 
case "转账":act="sjfunc/33.asp";break;case "清除":act="sjfunc/34.asp";break; 
case "修练":act="sjfunc/35.asp";break;case "暴豆":act="sjfunc/36.asp";break; 
case "哑穴":act="sjfunc/37.asp";break;case "好友":act="sjfunc/38.asp";break; 
case "怒吼":act="sjfunc/39.asp";break;case "心跳":act="sjfunc/40.asp";break; 
case "查ip":act="sjfunc/41.asp";break;case "奖励":act="sjfunc/42.asp";break; 
case "基金":act="sjfunc/43.asp";break;case "用卡":act="sjfunc/44.asp";break; 
case "今日话题":act="sjfunc/45.asp";break;case "招收弟子":act="sjfunc/46.asp";break; 
case "求婚":act="sjfunc/47.asp";break;case "幸运":act="sjfunc/48.asp";break; 
case "赌博":act="sjfunc/49.asp";break;case "下注":act="sjfunc/50.asp";break; 
case "对赌":act="sjfunc/51.asp";break;case "清屏":act="sjfunc/52.asp";break; 
case "情人":act="sjfunc/53.asp";break;case "分手":act="sjfunc/54.asp";break; 
case "单挑":act="sjfunc/55.asp";break;case "炸弹":act="sjfunc/56.asp";break; 
case "送花":act="sjfunc/57.asp";break;case "本派刑法":act="sjfunc/58.asp";break; 
case "本派罚款":act="sjfunc/59.asp";break;case "掌门令":act="sjfunc/60.asp";break; 
case "赌豆":act="sjfunc/61.asp";break;case "暂离":act="sjfunc/62.asp";break; 
case "回来":act="sjfunc/63.asp";break;case "离席":act="sjfunc/64.asp";break; 
case "双人赌博":act="sjfunc/65.asp";break;case "使用":act="sjfunc/66.asp";break; 
case "丢弃":act="sjfunc/67.asp";break;case "监禁ip":act="sjfunc/68.asp";break; 
case "封锁ip":act="sjfunc/69.asp";break;case "逐出":act="sjfunc/70.asp";break; 
case "标题":act="sjfunc/71.asp";break;case "配药":act="sjfunc/72.asp";break; 
case "宠物":act="sjfunc/73.asp";break;case "竞标":act="sjfunc/74.asp";break; 
case "通缉":act="sjfunc/75.asp";break;case "解除":act="sjfunc/75.asp";break; 
case "出家":act="sjfunc/76.asp";break;case "还俗":act="sjfunc/76.asp";break; 
case "休身养性":act="sjfunc/77.asp";break;case "念经":act="sjfunc/78.asp";break; 
case "如沐春风":act="sjfunc/78.asp";break;case "送豆点":act="sjfunc/79.asp";break; 
case "送金币":act="sjfunc/80.asp";break;case "乾坤一掷":act="sjfunc/81.asp";break; 
case "加入天网":act="sjfunc/82.asp";break;case "离开天网":act="sjfunc/82.asp";break; 
case "拐骗少女":act="sjfunc/83.asp";break;case "拐骗少男":act="sjfunc/84.asp";break; 
case "同归于尽":act="sjfunc/92.asp";break;case "练金属性":act="sjfunc/101.asp";break;  
case "练木属性":act="sjfunc/102.asp";break;case "练水属性":act="sjfunc/103.asp";break; 
case "练火属性":act="sjfunc/104.asp";break;case "练土属性":act="sjfunc/105.asp";break; 
case "赌金属性":act="sjfunc/187.asp";break;case "送歌":act="mtv/send.asp";break;
case "赌木属性":act="sjfunc/186.asp";break;case "赌水属性":act="sjfunc/185.asp";break;
case "赌火属性":act="sjfunc/184.asp";break;case "赌土属性":act="sjfunc/183.asp";break;
case "新人费":act="sjfunc/116.asp";break;case "天奔地裂":act="sjfunc/117.asp";break;
case "乞讨神术":act="sjfunc/165.asp";break;case "传送法力":act="sjfunc/164.asp";break;
case "存法力":act="sjfunc/137.asp";break;case "取法力":act="sjfunc/96.asp";break;
case "治病":act="sjfunc/209.asp";break;case "教武":act="sjfunc/208.asp";break;
case "求签":act="sjfunc/207.asp";break;case "发射":act="sjfunc/204.asp";break;
case "布施术":act="sjfunc/160.asp";break;
case "魔界咒语":act="sjfunc/159.asp";break;case "魅惑人间":act="sjfunc/158.asp";break;
case "迷魂大法":act="sjfunc/146.asp";break;case "爱神":act="sjfunc/85.asp";break;
case "颠倒":act="sjfunc/86.asp";break;case "移动":act="map/a1.asp";break;
case "血滴子":act="sjfunc/155.asp";break;case "轩辕":act="sjfunc/112.asp";break;
case "寻找法器":act="sjfunc/163.asp";break;case "执行":act="sjfunc/203.asp";break;
case "没收法器":act="sjfunc/94.asp";break;case "配制令牌":act="sjfunc/154.asp";break;
case "狼牙棒":act="sjfunc/157.asp";break;case "破天锥":act="sjfunc/156.asp";break;
case "使出":act="sjfunc/150.asp";break;case "传授轻功":act="map/1.asp";break;
case "轻功暂存":act="map/2.asp";break;case "提出轻功":act="map/3.asp";break;
case "讨取轻功":act="map/4.asp";break;case "寻找秘笈":act="map/5.asp";break;
case "宝物术":act="sjfunc/182.asp";break;case "炼金术":act="sjfunc/181.asp";break;
case "解毒术":act="sjfunc/180.asp";break;case "平安术":act="sjfunc/179.asp";break;
case "偷窃术":act="sjfunc/178.asp";break;case "摇钱术":act="sjfunc/177.asp";break;
case "蓝玛瑙":act="sjfunc/176.asp";break;case "帅哥令":act="sjfunc/175.asp";break;
case "美人令":act="sjfunc/174.asp";break;case "多情环":act="sjfunc/173.asp";break;
case "绝命钩":act="sjfunc/171.asp";break;case "粗体字":act="sjfunc/196.asp";break;
case "飞舞字":act="sjfunc/195.asp";break;case "按钮字":act="sjfunc/194.asp";break;
case "滚动按钮":act="sjfunc/193.asp";break;case "上下按钮":act="sjfunc/192.asp";break; 
case "修炼法力":act="sjfunc/118.asp";break;case "召唤":act="sjfunc/npc02.asp";break;
case "修炼轻功":act="sjfunc/119.asp";break;case "邀请跳舞":act="sjfunc/188.asp";break;
case "赌法力":act="sjfunc/190.asp";break;case "学成下山":act="sjfunc/202.asp";break; 
case "赌轻功":act="sjfunc/189.asp";break;case "走运气":act="sjfunc/139.asp";break;
case "传授道德":act="sjfunc/211.asp";break;case "传授魅力":act="sjfunc/210.asp";break;
case "倒夜香":act="sjfunc/206.asp";break;case "原子弹":act="sjfunc/100.asp";break;
case "官府奖励":act="sjfunc/122.asp";break;case "门派挑战":act="mp/mpopen.asp";break;
case "门派求和":act="mp/mpopen1.asp";break;case "门派大战":act="mp/mp1.asp";break;
case "打扑克":act="f2/dpk-ask.asp";break;case "杠牌":act="f2/DMJFP.ASP?action=6";break;
case "发牌":act="f2/dpkfp.asp";break;case "打麻将":act="f2/dmj-ask.asp";break;
case "出牌":act="f2/DMJFP.ASP?action=1";break;case "摸牌":act="f2/DMJFP.ASP?action=2";break;
case "问牌":act="f2/DMJFP.ASP?action=3";break;case "吃牌":act="f2/DMJFP.ASP?action=4";break;
case "碰牌":act="f2/DMJFP.ASP?action=5";break;case "偷金币":act="sjfunc/93.asp";break;
case "生育":act="sjfunc/97.asp";break;case "小孩":act="sjfunc/98.asp";break;
case "妙手回春":act="sjfunc/124.asp";break;case "寻找魔器":act="sjfunc/152.asp";break;
case "招财进宝":act="sjfunc/114.asp";break;case "养心大法":act="sjfunc/129.asp";break;
case "点石成金":act="sjfunc/115.asp";break;case "新雁南飞":act="sjfunc/125.asp";break;
case "新阴那山":act="sjfunc/130.asp";break;case "魔力钻石":act="sjfunc/162.asp";break;
case "生日蛋糕":act="sjfunc/144.ASP";break;case "百变神通":act="sjfunc/161.asp";break;
case "没收魔器":act="sjfunc/91.asp";break;case "配制宝石":act="sjfunc/151.asp";break;
case "发射子弹":act="sjfunc/145.ASP";break;case "绝情刀":act="sjfunc/143.ASP";break;
case "抢劫令":act="sjfunc/153.asp";break;case "字体魔法":act="sjfunc/191.asp";break;
case "移动魔法":act="sjfunc/191.asp";break;case "按钮魔法":act="sjfunc/191.asp";break;
case "发喜糖":act="sjfunc/121.asp";break;case "屏蔽":act="sjfunc/89.asp";break;
case "宝贝护主":act="sjfunc/197.asp";break;case "宝贝攻击":act="sjfunc/98.asp";break;
case "拿钱砸人":act="sjfunc/133.asp";break;case "扫黄行动":act="sjfunc/205.asp";break;
case "抢夺宝物":act="sjfunc/111.asp";break;case "振臂一呼":act="sjfunc/166.asp";break;
case "飞吻":act="sjfunc/127.asp";break;case "传授体力":act="sjfunc/88.asp";break;
case "传授武功":act="sjfunc/87.asp";break;case "邀请赏花":act="sjfunc/126.asp";break;
case "宠物自暴":act="sjfunc/90.asp";break;case "嫁衣神功":act="sjfunc/200.asp";break;
case "顿悟":act="sjfunc/199.asp";break;case "自杀":act="sjfunc/113.asp";break;
case "绿玛瑙":act="sjfunc/172.asp";break;case "红玛瑙":act="sjfunc/170.asp";break;
case "孔雀翎":act="sjfunc/169.asp";break;case "小孩护主":act="sjfunc/99.asp";break;
case "亲密接触":act="sjfunc/128.asp";break;case "暗然销魂":act="sjfunc/198.asp";break;
case "抢亲":act="sjfunc/120.asp";break;case "申领金币":act="sjfunc/131.asp";break;
case "新人福神":act="sjfunc/132.asp";break;case "宠物搏斗":act="sjfunc/201.asp";break;
case "抢劫银行":act="sjfunc/168.asp";break;case "抢劫金库":act="sjfunc/167.asp";break;
case "抓小偷":act="sjfunc/95.asp";break;case "册封花魁":act="sjfunc/134.asp";break;
case "巧嘴娃娃":act="sjfunc/135.asp";break;case "挑战掌门":act="sjfunc/136.asp";break;
case "查相同ip":act="sjfunc/138.asp";break;case "册封盟主":act="sjfunc/148.asp";break;
case "成为盟主":act="sjfunc/149.asp";break;case "盟主令":act="sjfunc/147.asp";break;
case "鸡毛掸子":act="sjfunc/140.asp";break;case "跪搓板":act="sjfunc/141.asp";break;
case "拧嘴":act="sjfunc/142.asp";break;case "加血":act="sjfunc/npc01.asp";break;
case "魔法石":act="sjfunc/013.asp";break;case "照顾新人":act="sjfunc/027.asp";break;
case "寻找鲜花":act="sjfunc/010.asp";break;case "寻找卡片":act="sjfunc/012.asp";break;
case "寻找金卡":act="sjfunc/011.asp";break;case "提点智力":act="sjfunc/018.asp";break;
case "追杀令":act="sjfunc/024.asp";break;case "杀气锐减":act="sjfunc/026.asp";break;
case "小吼吼":act="sjfunc/023.asp";break;case "大吼吼":act="sjfunc/025.asp";break;
case "暗器合成":act="sjfunc/015.asp";break;case "奖励徒弟":act="sjfunc/020.asp";break;
case "吃饭":act="sjfunc/034.asp";break;case "修习佛法":act="sjfunc/017.asp";break;
case "卡片合成":act="sjfunc/016.asp";break;case "炼魔法石":act="sjfunc/014.asp";break;
case "申领奖励":act="sjfunc/021.asp";break;case "生日":act="sjfunc/035.asp";break;
case "惩罚徒弟":act="sjfunc/019.asp";break;case "神灵降世":act="sjfunc/022.asp";break;
case "送金卡":act="sjfunc/033.asp";break;case "存金属性":act="sjfunc/106.asp";break;
case "存木属性":act="sjfunc/107.asp";break;case "存水属性":act="sjfunc/108.asp";break;
case "存火属性":act="sjfunc/109.asp";break;case "存土属性":act="sjfunc/110.asp";break;
case "取金属性":act="sjfunc/028.asp";break;case "取木属性":act="sjfunc/029.asp";break;
case "取水属性":act="sjfunc/030.asp";break;case "取火属性":act="sjfunc/031.asp";break;
case "取土属性":act="sjfunc/032.asp";break;case "寻找金币":act="sjfunc/037.asp";break;
case "介绍奖励":act="sjfunc/036.asp";break;case "邀请赏猪":act="sjfunc/212.asp";break;
case "新人奖励":act="sjfunc/038.asp";break;case "高级奖励":act="sjfunc/039.asp";break;
case "门派ip":act="sjfunc/040.asp";break;case "答题奖励":act="sjfunc/041.asp";break;
case "国家册封":act="sjfunc/guocf.asp";break;case "国家令":act="sjfunc/guoling.asp";break;
case "国家挑战":act="guo/guoopen.asp";break;case "国家大战":act="guo/guo1.asp";break;
case "国家求和":act="guo/guoopen1.asp";break;case "挑战君王":act="sjfunc/tzjw.asp";break;
case "国家招收":act="sjfunc/guozhao.asp";break;case "国库资金":act="sjfunc/guojj.asp";break;
case "国家赈灾":act="sjfunc/zhenzai.asp";break;case "唤回神兽":act="sjfunc/huanss.ASP";break;
case "神兽攻击":act="sjfunc/ssgj.ASP";break;case "欢迎新人":act="sjfunc/xinren.ASP";break;
case "人族任务":act="sjfunc/rzwc.asp";break;case "神族任务":act="sjfunc/szwc.asp";break;
case "魔族任务":act="sjfunc/mzwc.asp";break;case "修炼魔力":act="sjfunc/xiumo.asp";break;
case "修炼仙术":act="sjfunc/xiuxian.asp";break;case "修炼精神":act="sjfunc/xiujing.asp";break;
case "金币存库":act="sjfunc/jinkuc.asp";break;case "金库取币":act="sjfunc/jinkuq.asp";break;
case "神龙":act="sjfunc/shenlong.asp";break;case "邀请加入组":act="sjfunc/zu.asp";break;
case "观赏家园":act="sjfunc/gsjy.asp";break;case "邀请串门":act="sjfunc/seeroom.asp";break;
case "隐身":act="sjfunc/yinshen.ASP";break;case "册封状元":act="sjfunc/0134.asp";break;
case "邀请语聊":act="sjfunc/ipAV.asp";break;
default:alert('命令出错,无法执行')
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=saystemp;
this.f2.document.af.oldsays.value=saystemp;
this.f2.document.af.addsign.options[0].selected=true;
this.f2.document.af.tu.options[0].selected=true;
this.f2.document.af.sytemp.focus();
this.f2.document.af.sytemp.value='';return false;}}
else
if (nowmsg=="//"){act="say.asp";}else{act="say.asp"}
this.f2.document.af.action=act;
//this.f2.document.af.aqjhaction.value=act;
{this.f2.document.af.sy.value='';if(saystemp!=''){if((this.f2.document.af.oldsays.value==saystemp)&&(this.f2.document.af.oldtowho.value==this.f2.document.af.towho.value)){alert('内容不可重复！');
this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.select();return false;}this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=saystemp;this.f2.document.af.oldsays.value=saystemp;this.f2.document.af.addsign.options[0].selected=true;
this.f2.document.af.tu.options[0].selected=true;this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.value='';
ty=new Date();var nh=ty.getHours();var nm=ty.getMinutes();var ns=ty.getSeconds();var ct=(nh*3600)+(nm*60)+ns;
if(((ct-lst)<2.5)&&(ct>lst)){this.f2.af.sytemp.value=this.f2.af.oldsays.value;this.f2.af.oldsays.value='';return false;}
else{lst=ct;}this.f2.addOne(this.f2.document.af.sy.value);this.f2.startnosay();
this.f2.document.af.subsay.disabled=1;setTimeout('this.f2.document.af.subsay.disabled=0',3000);return true;}if ((this.f2.document.af.addsign.options[this.f2.document.af.addsign.selectedIndex].value=='0')||(this.f2.document.af.addsign.options[this.f2.document.af.addsign.selectedIndex].value=='')){alert('请输入发言或选择动作！');
this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.select();return false;}}}
function wdmess(b){this.mess.document.writeln(b);}
function mytitle(title){
this.title.document.open();
this.title.document.writeln("<html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=gb2312\"><\/head>");
this.title.document.writeln("<style type=\"text/css\">body{font-family:\"宋体\";color:blue;font-size:10.5pt;line-height:15pt;text-align:center}</style>");
this.title.document.writeln("<body bgcolor='E0E0F0' text=#FFFFFF TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0 oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false><div align=center style='font-size: 9pt;)'><font color=green>");
this.title.document.writeln(title+"</font></div></body></html>");
this.title.document.close();}
function DB(){this.mess.document.open();
wdmess("<html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=gb2312\"><\/head>");
wdmess("<style type=\"text/css\">body{background-image: URL(bbg1.jpg);font-family:\"宋体\";color:white;font-size:9pt;line-height:15pt;text-align:center}a{color:white;text-decoration:none;}a:hover{color:white;text-decoration:underline;}</style>");
wdmess("<body bgcolor=#A90005 TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0 oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false >");
var i=0;i=Math.ceil(Math.random()*(Banner.length-1));
wdmess("<marquee scrollamount=3 scrolldelay=10 onmouseover=this.stop(); onmouseout=this.start();>广告:" + Banner[i] + "<\/marquee><\Script Language=JavaScript>setTimeout('parent.DB()',100000);<\/script></body></html>");
this.mess.document.close();}
function tbclutch(){if(this.f2.document.af.tbclutch.value=='全屏'){this.f2.document.af.tbclutch.value='垂直';this.msgfrm.rows = "*";this.msgfrm.cols = "*";tbclu=false;}else{if(this.f2.document.af.tbclutch.value=='垂直'){this.f2.document.af.tbclutch.value='水平';this.msgfrm.cols = "*,*";this.msgfrm.rows = "*";tbclu=true;}else{this.f2.document.af.tbclutch.value='全屏';this.msgfrm.cols = "*";this.msgfrm.rows = "*,*";tbclu=true;}}this.f2.document.af.sytemp.focus();}
self.onerror=null;
var nullframe = '<HTML><BODY BGCOLOR=#000000 text=#ffffff style="font-size:10pt"><center><font color=yellow>欢迎来到<%=Application("aqjh_chatroomname")%></font><br><br>江湖正在从服务器读取资料, 请稍候^_^</center></BODY></HTML>';
</script>
</head>
<frameset cols="*,160" name=tbgn1 rows="*" framespacing="0" frameborder="no" border="3"  noresize>
<frameset rows="60,0,*,20,0,83,0" cols="*">
<frame src="toptupian.asp" scrolling="NO"  name="gg" marginwidth="3" marginheight="3" noresize>
<frame src="about:blank" name="mess" scrolling="no">
<frameset name=msgfrm rows="*,*" cols="*">
<frame src="javascript:parent.nullframe" name="f0" scrolling="AUTO" framespacing="1" marginheight="3" marginwidth="5" frameborder="yes" noresize>
<frame src="javascript:parent.nullframe" name="f1" scrolling="AUTO" framespacing="1" marginheight="3" marginwidth="5" frameborder="yes">
</frameset>
<frame src="about:blank" name="title" scrolling="no" noresize>
<frame src="about:blank" name="t" marginwidth="5" marginheight="5"  scrolling="NO" noresize>
<frame src="F2.asp" name="f2" scrolling="NO" marginwidth="3" marginheight="8" noresize>
<frame src="about:blank" name="d" scrolling="NO" noresize>
</frameset>
<frameset rows="0,0,22,*,0,0,104" cols="*" name="tbymd">
<frame src="about:blank" name="m">
<frame src="about:blank" name="ps">
<frame src="chang_room.asp" marginwidth="5" marginheight="5" scrolling="no" name="r" noresize>
<frame src="about:blank" marginwidth="5" marginheight="5" scrolling="auto" name="f3" noresize>
<frame src="menu.asp" scrolling="NO"  name="menu" marginwidth="3" marginheight="3" noresize>
<frame src="npc.asp" scrolling="NO"  name="npc" marginwidth="3" marginheight="3" noresize>
<frame src="F4.asp" scrolling="NO" name="CW_MENU" marginwidth="3" marginheight="3" noresize>
</frameset></frameset></html>