<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then 
	Session.Abandon
	Response.Write "<script language=JavaScript>{alert('提示：禁止使用代理注册！');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
if Application("sjjh_closedoor")="1" then
Response.Write "<script Language=Javascript>alert('欢迎您的光临！\n由于最近服务器不太稳定，所以站长决定暂时关闭社区注册功能\n现在不能进行注册，请过些时间再来！\n—— —— —— —— —— —— —— —— —— —— —— —— —— —— ——\n感谢您对本江湖的支持和厚爱！\n给您带来的不便，甚感抱歉！\nSEE YOU\n\n2003-11-10');window.parent.close();</script>"
Response.End
end if
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yusu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yusu=129 or yusu>192 or (yusu<2 and yusu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yusu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
'以下程序是禁止代理,可以防止世纪江湖攻击器攻击
if (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
name=server.HTMLEncode(trim(Request.form("name")))
sex=Request.form("sex")
zz=Request.form("zz")
psw=trim(Request.form("psw"))
pswc=trim(Request.Form("pswc"))
ask=trim(Request.form("ask"))
answer=trim(Request.form("answer"))
e_mail=LCase(Request.form("e_mail"))
oicq=trim(Request.form("oicq"))
jsr=trim(request.form("jsr"))
face=request("face")
if Request.form("face")=00 then
	Response.Write "<script language=JavaScript>{alert('提示：您还没有选择头像呢！这可是您在聊天室的形象哦！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if not isnumeric(oicq) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&oicq&"]输入错误，OICQ请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');</script>"
		Response.End
end if
next
name=trim(name)
name=server.HTMLEncode(name)
if jsr=name then response.redirect "../error.asp?id=57"
if name="无" or name="未定" then response.redirect "../error.asp?id=130"
if chuser(name) then Response.Redirect "../error.asp?id=120"
if chuser(jsr) then Response.Redirect "../error.asp?id=60"
if len(oicq)<4 or len(oicq)>=15 then Response.Redirect "../error.asp?id=50"
if instr(e_mail,"@")=0 then Response.Redirect "../error.asp?id=51"
if len(pswc)<5 then Response.Redirect "../error.asp?id=52"
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
if Instr(name,"主席")>0 or Instr(name,"交易人名")>0 or Instr(name,"管理")>0 or Instr(name,"朋友名字")>0 or Instr(name,Application("sjjh_automanname"))>0 or Instr(name,"回首当年")>0 or Instr(name,"伊然")>0 or Instr(name,"站长")>0 or Instr(name,"网管")>0 or Instr(name,"时代")>0 or Instr(name,"江湖")>0 or Instr(name,"中国")>0 or Instr(name,"严")>0 or Instr(name,"妈")>0 or Instr(name,"爸")>0 or Instr(name,"大家")>0 or Instr(name,"操")>0 or Instr(name,"黑客")>0 or Instr(name,"嘿客")>0 then Response.Redirect "../error.asp?id=130"
if pswc<>psw then Response.Redirect "../error.asp?id=166"
if trim(request.form("Name"))="" or trim(request.form("psw"))="" or trim(Request.Form("e_mail"))="" or trim(request.form("oicq"))="" then Response.Redirect "../error.asp?id=56"
if trim(request.form("Name"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=129"
if left(name,3)="%20" OR InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name,"　")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 then Response.Redirect "../error.asp?id=120"
if left(jsr,3)="%20" OR InStr(jsr,"=")<>0 or InStr(jsr,"`")<>0 or InStr(jsr,"'")<>0 or InStr(jsr," ")<>0 or InStr(jsr,"　")<>0 or InStr(jsr,"'")<>0 or InStr(jsr,chr(34))<>0 or InStr(jsr,"\")<>0 or InStr(jsr,",")<>0 or InStr(jsr,"<")<>0 or InStr(jsr,">")<>0 then Response.Redirect "../error.asp?id=120"
if len(jsr)=0 then jsr="无"
psw=md5(psw)
zaohui=md5(ask&answer)
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
if y=2 And r>28 Then
	r=28
Else
	if r=31 Then	
		r=30 
	End if
end if
if y<10 then
huiqi=n & "-" & y+3 & "-" & r
else
huiqi=n+1 & "-" & y+3-12 & "-" & r
end if
userip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE regip=" & SqlStr(userip) & " and DateDiff('s',regtime,#" & sj & "#)<=150",conn,1,1
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=67"
			response.end
end if
rs.close
rs.open "SELECT * FROM 用户 WHERE 姓名='" & name & "'",conn,1,1
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
rs.open "select top 1 * FROM 用户 WHERE DateDiff('d',lasttime,now())>90 and grade<5",conn,1,2
if rs.eof or rs.bof then
rs.AddNew
end if 
rs("姓名")=name
rs("性别")=sex
rs("年龄")=10
rs("名单头像")=face
rs("头像")=touxiang
rs("密码")=psw
rs("门派")="游侠"
rs("门派基金")=0
rs("身份")="弟子"
rs("职业")="采冰"
rs("师傅")="无"
rs("师傅交钱")="无"
rs("登录")=now()
rs("信箱")=e_mail
rs("oicq")=oicq
rs("魅力")=300
rs("道德")=300
rs("赢钱")=0
rs("帖子数")=0
rs("赌次数")=0
rs("赢次数")=0
rs("武功")=250000
rs("武功加")=0
rs("内力")=120000
rs("内力加")=0
rs("体力")=300000
rs("体力加")=0
rs("攻击")=0
rs("防御")=0
rs("知质")=100
rs("银两")=1000000
rs("状态")="正常"
rs("grade")=1
rs("等级")=10
rs("times")=1
rs("regtime")=now()
rs("regip")=userip
rs("lasttime")=now()
rs("lastip")=userip
rs("allvalue")=5000
rs("mvalue")=0
rs("领钱")=now()
rs("保留")="保留"
rs("保留1")="保留"
rs("保留2")="保留"
rs("存款")=10000
rs("结算日期")=date()
rs("地区")="未知"
rs("签名")="WWW.happyjh.com"
rs("会员等级")=1
rs("会员日期")=huiqi
rs("会员金卡")=0
rs("宝物")="无"
rs("宝物修练")=0
rs("介绍人")=jsr
rs("泡豆点数")=0
rs("暴豆时间")=now()-1
rs("操作时间")=now()
rs("好友名单")="|回首当年|"
rs("结婚次数")=0
rs("结婚记念日")=date()
rs("情人")="无"
rs("配偶")="无"
rs("事件原因")="无"
rs("保护")=true
rs("通缉")=false
rs("杀人数")=0
rs("总杀人")=0
rs("死亡时间")=now()
rs("金币")=0
rs("会员")=false
rs("会员结束")=date()
rs("充值卡")=0
rs("银币")=0
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
rs("二号情人")="无"
rs("三号情人")="无"
rs("sl")="无"
rs("slsj")=now()
rs("cw")=""
rs("w1")="a"
rs("w2")="a"
rs("w3")="a"
rs("w4")="a"
rs("w5")="a"
rs("w6")="a"
rs("w7")="a"
rs("w8")="a"
rs("w9")="a"
rs("w10")="a"
rs("z1")=""
rs("z2")=""
rs("z3")=""
rs("z4")=""
rs("z5")=""
rs("z6")=""
rs("hua")=""
rs.Update
rs.close
rs.open "SELECT id,姓名,性别,名单头像,介绍人,regtime FROM 用户 WHERE 姓名='" & name & "'",conn,1,1
id=rs("id")
xinren=rs("姓名")
sex=rs("性别")
touxiang=rs("名单头像")
jsr=rs("介绍人")
regsj=rs("regtime")
rs.close
set rs=nothing
conn.close
set conn=nothing
Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
End Function
says="<bgsound src=wav/Global.wav loop=1><img src=../yamen/xinr.gif><font color=red>〖新人加入〗</font><font color=green>大侠</font><img src=../ico/"& touxiang &"-2.gif width=32 height=32><font color=#000000>[</font><font color=#cc0000>"& xinren &"</font>("& sex &")<font color=#000000>]</font><font color=green>ID</font><font color=#000000>[</font><font color=#cc0000>"& id &"</font><font color=#000000>]</font><font color=green>加入了快乐江湖总站，大家对新人要多照顾啊！</font><font color=green>介绍人</font><font color=#000000>[</font><font color=#cc0000>"& jsr &"</font><font color=#000000>]</font>(<font color=#ff00ff>" & regsj & "</font>)"			'聊天数据
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('恭喜{" & name & "}，您已经成功加入本社区，请您登陆！\n以下信息请牢记，登陆后应及时申请密码保护！\n—— —— —— —— —— —— —— —— —— —— —— ——\n★ID：" & id & "\n姓名：" & name & "\n密码：" & pswc & "\n如果连续两个月未登陆\n您的ID将被新用户取代！');window.parent.close();</script>"
Response.End
%>
nse.End
%>