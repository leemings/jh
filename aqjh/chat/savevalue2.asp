<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../paodian.asp"-->
<!--#include file="../const3.asp"-->
<!--#include file="../chk.asp"-->
<!--#include file="sjfunc/func.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
nowinroom=session("nowinroom")
id=LCase(trim(request.querystring("id")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
ChkPost()
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if aqjh_disproxy=1 then Chkproxy()
%>
<script language="JavaScript"> 
//if(window.top==window.self){var i=1;while (i<=50){window.alert("你想作什么呀，黑我？这里是不行的，去别处玩去吧！哈！慢慢点50次！！");i=i+1;}top.location.href="../exit.asp"}
</script>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_yjdh=1 then Chkyjdh()
inroom=session("nowinroom")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
%>
<html>
<head>
<META http-equiv='content-type' content='text/html; charset=gb2312'>
<title>保存经验值</title>
<link rel="stylesheet" href="READONLY/STYLE.CSS">
<style>
body{
CURSOR: url('boy.cur');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" topmargin="5" text="#FFFFFF">
<%
useronlinename=Application("aqjh_useronlinename"&inroom)
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
ip=Request.ServerVariables("REMOTE_ADDR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
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
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
mycd=DateDiff("n",Session("aqjh_savetime"),now())
addvalue=mycd
addvalue1=mycd*aqjh_paofen
Session("aqjh_savetime")=now()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE  姓名='" & aqjh_name &"'",conn,1,3
mysl=clng(DateDiff("n",now(),rs("slsj")))
mysls="["&rs("sl")&"]"
olddj=rs("等级")
jsr=rs("介绍人")
zhuans=rs("转生")
if jsr<>"无"  then
jiang=1
else
jiang=0
end if
if olddj < 90 and rs("转生")<1 then
aqjh_paofencd=(aqjh_paofencd*2.5)
else
aqjh_paofencd=aqjh_paofencd
end if
if mysl>0 then
	Select Case rs("sl")
	case "福神"
		aqjh_paofencd=aqjh_paofencd*1.5
	case "冰神"
		aqjh_paofencd=aqjh_paofencd*2.5
	case "封神"
		aqjh_paofencd=aqjh_paofencd*2
	case "财神"
		aqjh_paofenyin=aqjh_paofenyin*5	
	case "衰神"
		aqjh_paofencd=int(aqjh_paofencd/2)
	case "死神"
		aqjh_paofen=-(aqjh_paofen*10)
	case "穷神"
		aqjh_paofenyin=-(aqjh_paofenyin*5)
	end select
else
	rs("sl")="无"
end if

if rs("会员")=true and clng(DateDiff("d",now(),rs("会员结束")))>0 then
	pd=paodian
	pdstr="泡点制会员"
else
	pd=1
	pdstr="非泡点会员"
end if
sf=rs("师傅")
hydj=rs("会员等级")
zddj=rs("等级")
dj1=rs("等级")
sfwg=1
if sf<>"" and sf<>"无" then
	if Instr(LCase(Application("aqjh_useronlinename"&inroom))," "&LCase(sf)&" ")=0 then
		sfwg=1
	else
		sfwg=2
	end if
end if
jhhy=1
Select Case hydj
case 1
	jhhy=hy1
case 2
	jhhy=hy2
case 3
	jhhy=hy3
case 4
	jhhy=hy4
case 5
	jhhy=hy5
case 6
	jhhy=hy6
case 7
	jhhy=hy7
case 8
	jhhy=hy8
case else
	jhhy=1
end select
rs("内力")=rs("内力")+int(addvalue1*sfwg*jhhy)
rs("银两")=rs("银两")+int(addvalue1*jhhy*aqjh_paofenyin)
rs("武功")=rs("武功")+int(addvalue1/3*sfwg*jhhy)
rs("攻击")=rs("攻击")+int(addvalue1/3*jhhy)
rs("防御")=rs("防御")+int(addvalue1/3*jhhy)
rs("法力")=rs("法力")+int(addvalue1/3*jhhy)
rs("轻功")=rs("轻功")+int(addvalue1/3*jhhy)
rs("粪库")=rs("粪库")+int(addvalue1/16*jhhy)
rs("元气")=rs("元气")+int(addvalue1*sfwg*jhhy)
rs("金")=rs("金")+int(addvalue1/40)
rs("木")=rs("木")+int(addvalue1/40)
rs("水")=rs("水")+int(addvalue1/40)
rs("火")=rs("火")+int(addvalue1/40)
rs("土")=rs("土")+int(addvalue1/40)
bl=rs("配偶")
rs("allvalue")=rs("allvalue")+int(addvalue*jhhy*aqjh_paofencd)*pd
rs("泡豆点数")=rs("泡豆点数")+addvalue
prevtime=CDate(rs("lasttime"))
if DateDiff("m",prevtime,now())=0 then
	rs("mvalue")=rs("mvalue")+int(addvalue*jhhy*aqjh_paofencd)*pd
else
	rs("mvalue")=addvalue*pd
end if
rs("lasttime")=sj
rs("lastip")=ip
rs.Update
aqjh_value=rs("allvalue")
aqjh_mvalue=rs("mvalue")
dengji=int(sqr(aqjh_value/50))
rs("等级")=dengji
rs.Update
Session("aqjh_grade")=rs("grade")
Session("aqjh_jhdj")=rs("等级")
zddj=rs("等级")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 粪库,状态,转生,性别 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("状态")="死" then
call boot(aqjh_name,"无") 
Response.Write "<script language=JavaScript>{alert('你可能在舞剑或孤岛中被打死!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("粪库")>=50000 then
conn.execute "update 用户 set 状态='死',事件原因='被自己粪淹死|"&fn1&"' where 姓名='" & aqjh_name & "'"
call boot(aqjh_name,aqjh_name&"被"&aqjh_name&"的粪淹死了") 
Response.Write "<script language=JavaScript>{alert('你粪已达50000点被自己粪淹死了!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

'判断升级
if zddj>dj1 then
if jiang=1 then
if zhuans>1 then 
conn.execute "update 用户 set 金币=金币+2,智力=智力+100  where 姓名='" & jsr & "' "
mess="<br><font color=blue><font color=red>["&aqjh_name&"]</font>达到系统奖励条件$<font color=red>转生2次以上</font>$,因此介绍人<font color=red>["&jsr&"]</font>在<font color=red>["&aqjh_name&"]</font>每次升级的时候得到系统奖励<font color=brown>金币</font>2个,<font color=brown>智力</font>100点！！</font>"
else
mess="<br><font color=blue><font color=red>["&aqjh_name&"]</font>希望你多多拉人来快乐江湖，系统有奖励的哦!</font>"
end if
else 
mess="<br><font color=blue><font color=red>["&aqjh_name&"]</font>没有介绍人，多拉人来，自己有奖励的哦!</font>"
end if 
if rs("性别")="男" then
says="<img src=img/up.gif><bgsound src=wav/system.wav loop=1><font color=green>〖升级消息〗恭喜</font>大帅哥<font color=red>["& aqjh_name &"]</font><font color=green>经过坚持不懈的努力,又升级了！<font color=red><b><img src=img/sjl.gif>"&dj1&"→<img src=img/sjl.gif>"&zddj&"</b></font>,<font color=blue>∮转生"&zhuans&"次∮</font>,愿你越来越帅，越来越厉害，吸引更多美女哦，帅哥加油哦......</font>"&mess
else
says="<img src=img/up1.gif><bgsound src=wav/system.wav loop=1><font color=green>〖升级消息〗恭喜</font>大美女<font color=red>["& aqjh_name &"]</font><font color=green>经过坚持不懈的努力,又升级了！<font color=red><b><img src=img/sjl.gif>"&dj1&"→<img src=img/sjl.gif>"&zddj&"</b></font>,<font color=blue>∮转生"&zhuans&"次∮</font>,祝大美女越来越漂亮，嘿嘿，美女加油哦......</font>"&mess
end if
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& inroom & ");<"&"/script>"
addmsg saystr
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
end if
If olddj < 90 and rs("转生")<1 Then
Response.Write "等级低于90级,0级转生新人,泡点翻2.5倍!<br>"
end if
%>
<table width="100%" border="1" cellspacing="0" cellpadding="5" align="center" height="267" bordercolor="#F7DEAC">
 <tr valign="top">
<td height="260" style="filter:dropshadow(color=black, offx=1,offy=1);">
<p align="center"> <%=aqjh_name%><br>当 前 状 态</p>
<div align="center"><font color=yellow>
<%if sf<>"" and sf<>"无" and sfwg=2 then%> 
	师傅:<font color=red><%=sf%></font>在场武功、内力大增！<br> 
<%end if%>
<%if hydj>0  then%>您是<%=hydj%>级会员<br>泡点为<%=jhhy%>倍<%else%>非会员泡点不翻倍<%end if%> 
<br><font color=red><%if olddj < 90 and rs("转生")<1 then%>等级小于90,0级转生人,泡点为2.5倍,加油哦<%end if%>
</font><br><%=pdstr%>
<%if mysl>0 then%><HR><%=mysls%><br>还有<%=mysl%>分钟<HR><%end if%> 
</div></font>
<p align="left"> 
<%sjjf=(zddj+1)*(zddj+1)*50-aqjh_value%>
时间:<%=mycd%>分钟<br> 
管理:<%=aqjh_grade%> 级<br> 
战斗:<%=int(sqr((aqjh_value/50)))%> 级<br> 
月积:<%=aqjh_mvalue%><br>
总积:<%=aqjh_value%><br>
升差:<font color=red><%=sjjf%></font>点<br>
泡点:<font color=red><%=int(addvalue*aqjh_paofencd*jhhy)*pd%></font>点<br>
<br>
你共增加<br>
内力:<%if sfwg=2 then%><font color=red><%=int(addvalue1*sfwg*jhhy)%></font>点<br><%else%><%=int(addvalue1*sfwg*jhhy)%>点<br><%end if%> 
银子:<%=addvalue1*aqjh_paofenyin%>两<br> 
武功:<%if sfwg=2 then%><font color=red><%=int(addvalue1/2*sfwg*jhhy)%></font>点<br><%else%><%=int(addvalue1/2*sfwg*jhhy)%>点<br><%end if%> 
攻击:<%=int(addvalue1/3*jhhy)%>点<br> 
防御:<%=int(addvalue1/3*jhhy)%>点<br> 
法力:<%=int(addvalue1/3*jhhy)%>点<br> 
轻功:<%=int(addvalue1/3*jhhy)%>点<br>
粪库:<%=int(addvalue1/16*jhhy)%>点<br> 
金:<font color=red><%=int(addvalue1/40)%></font>点<br>
木:<font color=red><%=int(addvalue1/40)%></font>点<br>
水:<font color=red><%=int(addvalue1/40)%></font>点<br>
火:<font color=red><%=int(addvalue1/40)%></font>点<br>
土:<font color=red><%=int(addvalue1/40)%></font>点<br>
<br>
<%if bl<>"无" and bl<>"" then
peioujia=int(addvalue1/4)
conn.execute "update 用户 set 内力=内力+" & addvalue1/2 & ",银两=银两+" & peioujia & ",武功=武功+" & peioujia & ",攻击=攻击+" & peioujia & ",法力=法力+"  & peioujia & ",轻功=轻功+"  & peioujia & ",防御=防御+" & peioujia & " where 姓名='" & bl & "'"%>
伴侣<font color=red> <%=bl%></font><br> 
内力:<%=int(addvalue1/4)%>点<br> 
银两:<%=int(addvalue1/4)%>两<br> 
武功:<%=int(addvalue1/4)%>点<br> 
攻击:<%=int(addvalue1/4)%>点<br> 
防御:<%=int(addvalue1/4)%>点<br> 
元气:<%=int(addvalue1/2*jhhy)%>点<br>
法力:<%=int(addvalue1/4)%>点<br> 
轻功:<%=int(addvalue1/4)%>点<br> 
<%end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</td></tr></table></html>