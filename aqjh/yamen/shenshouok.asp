<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../const1.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
name=LCase(trim(Request.form("name")))
sex=trim(Request.form("sex"))
psw=trim(Request.form("psw"))
pswc=trim(Request.Form("pswc"))
psw2=trim(Request.form("psw2"))
pswc2=trim(Request.Form("pswc2"))
oicq=trim(Request.form("oicq"))
e_mail=LCase(Request.form("e_mail"))
jsr=trim(request.form("jsr"))
face=request("face")
face="../ico/"&face&"-2.gif"
regjm1=trim(request("regjm1"))
regjm2=trim(request("regjm2"))
if not isnumeric(regjm1) or regjm1<>regjm2 then
	Response.Write "<script language=JavaScript>{alert('提示：认证码输入错误，系统拒绝你注册！');window.close();}</script>"
	Response.End 
end if
if not isnumeric(oicq) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&oicq&"]输入错误，oicq请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
name=trim(name)
name=server.HTMLEncode(name)
if jsr=name then response.redirect "../error.asp?id=57"
if name="无" or name="未定" then response.redirect "../error.asp?id=130"
if chuser(name) then Response.Redirect "../error.asp?id=120"
if chuser(jsr) then Response.Redirect "../error.asp?id=60"
if len(oicq)<4 or len(oicq)>=10 then Response.Redirect "../error.asp?id=50"
if instr(e_mail,"@")=0 then Response.Redirect "../error.asp?id=51"
if len(pswc)<6 then Response.Redirect "../error.asp?id=52"
if len(pswc2)<6 then Response.Redirect "../error.asp?id=52"
for i=1 to len(name)
usernamechr=mid(name,i,1)
if asc(usernamechr)>0 then  Response.Redirect "../error.asp?id=120"
next
for i=1 to len(jsr)
usernamechr=mid(jsr,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=60"
next
if instr(name,"or")<>0 or instr(sex,"or")<>0 or instr(psw,"or")<>0 or instr(pswc,"or")<>0 or instr(email,"or")<>0 or instr(oicq,"or")<>0 or instr(ask,"or")<>0 or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if instr(name,"=")<>0 or instr(sex,"=")<>0 or instr(psw,"=")<>0 or instr(pswc,"=")<>0 or instr(email,"=")<>0 or instr(oicq,"=")<>0 or instr(ask,"=")<>0  or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
'if Instr(name,"交易人名")>0 or Instr(name,"公")>0 or Instr(name,"公告")>0 or Instr(name,"公告江")>0 or Instr(name,"奥星")>0 or Instr(name,"剑心")>0 or Instr(name,"八")>0 or Instr(name,"季节")>0 or Instr(name,"风")>0 or Instr(name,"络")>0 or Instr(name,"天度网")>0 or Instr(name,"天度网")>0 or Instr(name,"园网络")>0 or Instr(name,"官府")>0 or Instr(name,"管理")>0 or Instr(name,"朋友名字")>0 or Instr(name,Application("aqjh_automanname"))>0 or Instr(name,"天度网络")>0 or Instr(name,"站长")>0 or Instr(name,"网管")>0 or Instr(name,"时代")>0 or Instr(name,"严")>0 or Instr(name,"妈")>0 or Instr(name,"爸")>0 or Instr(name,"大家")>0 or Instr(name,"操")>0 then Response.Redirect "../error.asp?id=130"
if pswc<>psw then Response.Redirect "../error.asp?id=166"
if pswc2<>psw2 then Response.Redirect "../error.asp?id=166"
if trim(request.form("Name"))="" or trim(request.form("psw"))="" or trim(Request.Form("e_mail"))="" or trim(request.form("oicq"))="" then Response.Redirect "../error.asp?id=56"
if trim(request.form("Name"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=129"
if left(name,3)="%20" OR InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name,"　")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 then Response.Redirect "../error.asp?id=120"
if left(jsr,3)="%20" OR InStr(jsr,"=")<>0 or InStr(jsr,"`")<>0 or InStr(jsr,"'")<>0 or InStr(jsr," ")<>0 or InStr(jsr,"　")<>0 or InStr(jsr,"'")<>0 or InStr(jsr,chr(34))<>0 or InStr(jsr,"\")<>0 or InStr(jsr,",")<>0 or InStr(jsr,"<")<>0 or InStr(jsr,">")<>0 then Response.Redirect "../error.asp?id=120"
if len(jsr)=0 then
	jsr="无"
end if
psw=md5(psw)
psw2=md5(psw2)
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
userip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
s35=0
s36=10000
s37=0
s38=360

rs.open "SELECT * FROM 用户 WHERE 姓名='" & name & "'",conn
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=62"
			response.end
end if
rs.close
if sex="男" then
	touxiang="../jhimg/boy00.gif"
else
	touxiang="../jhimg/girl00.gif"
end if
rs.open "select top 1 * FROM 用户 WHERE CDate(lasttime)<(now()-"& s38 &")",conn,1,2
if rs.eof or rs.bof then
rs.AddNew
end if
conn.Execute ("update 用户 set 召唤兽1='"&name&"' where 姓名='"&aqjh_name&"'")

'rs.Open "SELECT * FROM 用户",conn,1,2
'Select count(*) As 数量 from 用户 where 
rs("姓名")=name
rs("性别")=sex
rs("年龄")=10
rs("名单头像")=face
rs("头像")=touxiang
rs("密码")=psw
rs("第二密码")=psw2
rs("门派")="游侠"
rs("门派基金")=0
rs("身份")="弟子"
rs("职业")="采冰"
rs("师傅")="无"
rs("师傅交钱")="无"
rs("登录")=now()
rs("信箱")=e_mail
rs("oicq")=oicq
rs("魅力")=reg_ml
rs("道德")=reg_dd
rs("赢钱")=0
rs("帖子数")=0
rs("赌次数")=0
rs("赢次数")=0
rs("武功")=reg_wg
rs("武功加")=0
rs("内力")=reg_nl
rs("内力加")=0
rs("体力")=reg_tl
rs("体力加")=0
rs("攻击")=reg_gj
rs("防御")=reg_fy
rs("知质")=reg_zz
rs("法力")=reg_fl
rs("法力加")=0
rs("轻功")=reg_qg
rs("轻功加")=0
rs("银两")=reg_yl
rs("状态")="正常"
rs("grade")=1
rs("等级")=reg_dj
rs("times")=1
rs("regtime")=now()
rs("regip")=userip
rs("lasttime")=now()
rs("lastip")=userip
rs("allvalue")=reg_dj*reg_dj*50
rs("mvalue")=0
rs("领钱")=now()
rs("保留")="保留"
rs("保留1")="保留"
rs("保留2")="保留"
rs("存款")=reg_ck
rs("结算日期")=date()
rs("地区")="未知"
rs("签名")="爱情江湖网"
rs("会员等级")=reg_hydj
rs("会员日期")=date()+reg_hydate
rs("会员金卡")=reg_hyjk
rs("杀人数")=0
rs("宝物")="无"
rs("宝物修练")=0
rs("介绍人")=jsr
rs("泡豆点数")=0
rs("暴豆时间")=now()
rs("操作时间")=now()
rs("好友名单")="|"&Application("aqjh_user")&"|"
rs("结婚次数")=0
rs("结婚记念日")=date()
rs("情人")="无"
rs("配偶")="无"
rs("事件原因")="无"
rs("保护")=true
rs("通缉")=false
rs("杀人数")=0
rs("总杀人")=0
rs("死亡时间")=now()-1
rs("金币")=reg_jb
rs("会员")=false
rs("会员结束")=date()
rs("金")=0
rs("木")=0
rs("水")=0
rs("火")=0
rs("土")=0
rs("金加")=0
rs("木加")=0
rs("水加")=0
rs("火加")=0
rs("土加")=0
rs("sl")="无"
rs("slsj")=now()
rs("cw")=""
rs("w1")="无"
rs("w2")="无"
rs("w3")="无"
rs("w4")="无"
rs("w5")=reg_hycard
rs("w6")="无"
rs("w7")="无"
rs("w8")="无"
rs("w9")="无"
rs("w10")="无"
rs("z1")=" "
rs("z2")=" "
rs("z3")=" "
rs("z4")=" "
rs("z5")=" "
rs("z6")=" "
rs("转生")=0
rs("花魁")=0
rs("攻击加")=0
rs("防御加")=0
rs("属性")="神兽"
rs.Update
rs.close
rs.open "SELECT id FROM 用户 WHERE 姓名='" & name & "'",conn
id=rs("id")
rs.close
set rs=nothing
conn.close
set conn=nothing
Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
End Function

says="<font color=red>〖神兽诞生〗</font><font color=green>大侠</font><font color=#0000ff>[</font><font color=#0000ff>"& aqjh_name &"</font><font color=#000000>]</font>经过2次转生,参悟生死之道.终于召唤出他自己的神兽   [ "&name&" ]."   '聊天数据
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu1(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#FFFFFF" background="../hyzq/bg.gif">
<table border=1 align=center width=400 cellpadding="10" cellspacing="13" height="293" background="../images/b2.gif">
<tr valign="top" bgcolor="0088FF">
<td height="226">
<div align="center">
<p><font size="3"><b>您成功建立神兽</b></font><br>
<p><br>
<br>
<br>
</div>
<table width=100%>
<tr>
<td>
<p align=center style='font-size:14;color:red'><font color="#FF6600">注册神兽ID<font color="#0000FF">:<%=id%></font></font><br>
<font color="#FF6600">注册神兽名<font color="#0000FF">:<%=name%></font><br>
注册神兽头像：<img src="../ico/<%=face%>"><br>
　</font></p>
<p align=center>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
</html>

