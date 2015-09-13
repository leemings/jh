<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../const1.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("aqjh_disproxy")=1 then
if (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
end if
'禁止站外提交数据
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
	Response.Write "<script language=JavaScript>{alert('提示：禁止站外提交数据！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if Application("aqjh_closedoor")="1" then
Response.Write "<script Language=Javascript>alert('欢迎您的光临！\n由于最近服务器不太稳定，所以站长决定暂时关闭社区注册功能\n现在不能进行注册，请过些时间再来！\n—— —— —— —— —— —— —— —— —— —— —— —— —— —— ——\n感谢您对本江湖的支持和厚爱！\n给您带来的不便，甚感抱歉！\nSEE YOU\n\n2004-4-7');window.parent.close();</script>"
Response.End
end if
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
sex=Request.form("sex")
country=Request.Form("country")
psw=trim(Request.form("psw"))
pswc=trim(Request.Form("pswc"))
psw2=trim(Request.Form("psw2"))
oicq=trim(Request.form("oicq"))
sex=Request.form("sex")
e_mail=LCase(Request.form("e_mail"))
jsr=trim(request.form("jsr"))
jsr1=jsr
face=request("face")
face="../ico/"&face&"-2.gif"

if not isnumeric(oicq) then 
	Response.Write "<script language=JavaScript>{alert('提示：输入错误，oicq请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
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



if len(pswc)<5 or len(psw2)<5 then Response.Redirect "../error.asp?id=52"
for i=1 to len(name)
usernamechr=mid(name,i,1)
if asc(usernamechr)>0 then  Response.Redirect "../error.asp?id=120"
next
for i=1 to len(jsr)
usernamechr=mid(jsr,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=60"
next


 if instr(name,"or")<>0 or instr(sex,"or")<>0 or  instr(psw,"or")<>0  or instr(pswc,"or")<>0 or instr(email,"or")<>0 or instr(oicq,"or")<>0 or instr(ask,"or")<>0 or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if instr(name,"=")<>0 or instr(sex,"=")<>0  or  instr(psw,"=")<>0 or instr(pswc,"=")<>0 or instr(email,"=")<>0 or instr(oicq,"=")<>0 or instr(ask,"=")<>0  or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if Instr(name,"交易人名")>0 or Instr(name,"管理")>0 or Instr(name,"朋友名字")>0 or Instr(name,Application("aqjh_automanname"))>0 or Instr(name,"双鱼娱乐")>0 or Instr(name,"站长")>0 or Instr(name,"网管")>0 or Instr(name,"时代")>0 or Instr(name,"粟")>0 or Instr(name,"妈")>0 or Instr(name,"爸")>0 or Instr(name,"大家")>0 or Instr(name,"操")>0 then Response.Redirect "../error.asp?id=130"
if pswc<>psw then Response.Redirect "../error.asp?id=166"
if trim(request.form("Name"))="" or trim(request.form("psw"))="" or trim(Request.Form("e_mail"))="" or trim(request.form("oicq"))="" or trim(request.form("psw2"))="" then Response.Redirect "../error.asp?id=56"
if trim(request.form("Name"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=129"
if trim(request.form("psw2"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=485"
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
rs.open "SELECT * FROM 用户 WHERE regip=" & SqlStr(userip) & " and DateDiff('s',regtime,#" & sj & "#)<=10",conn
'If not(Rs.Bof OR Rs.Eof) Then
'			rs.close
'			set rs=nothing
'			conn.close
'			set conn=nothing
'			Response.Redirect "../error.asp?id=67"
'			response.end
'end if
rs.close
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
rs.open "SELECT n姓名 from npc WHERE n姓名='" & name & "'",conn,1,1
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=484"
			response.end
end if
rs.close



if jsr1<>"" then
rs.open "SELECT * FROM 用户 WHERE 姓名='" & jsr & "'",conn
If Rs.Eof Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=58"
			response.end
end if
rs.close
end if
if sex="男" then
	touxiang="../jhimg/boy00.gif"
else
	touxiang="../jhimg/girl00.gif"
end if
rs.open "select top 1 * FROM 用户 WHERE DateDiff('d',lasttime,now())>12000000 and grade<5",conn,1,3
if rs.eof or rs.bof then
rs.AddNew
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
rs("国家")=country
rs("职位")="百姓"
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
rs("签名")="www.7758530.com-快乐江湖网"
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
rs("离派")=0
rs("佛法")=0
rs("智力")=0
rs("威望")=0
rs("精神")=0
rs("魔力")=0
rs("仙术")=0
rs("黄金")=0
rs("种族")="无"
rs("进化")="初学忍者"
rs.Update
else
  Response.Redirect "../error.asp?id=67"
  response.end
end if
rs.close
rs.open "SELECT * FROM 用户 WHERE 姓名='" & name & "'",conn
id=rs("id")
xinren=rs("姓名")
sex=rs("性别")
touxiang=rs("名单头像")
jsr=rs("介绍人")
regsj=rs("regtime")
rs.close
set rs=nothing
Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
End Function
says="<bgsound src=wav/Global.wav loop=1><img src=../yamen/xinr.gif><font color=red>〖新人加入〗身居[</font><font color=#cc0000>"& country &"</font><font color=#000000>]的</font><font color=green>大侠</font><img src="& touxiang &"><font color=#000000>[</font><font color=#cc0000>"& xinren &"</font>("& sex &")<font color=#000000>]</font><font color=green>ID</font><font color=#000000>[</font><font color=#cc0000>"& id &"</font><font color=#000000>]</font><font color=green>加入江湖，大家对新人要多照顾啊！注册QQ:</font><font color=#000000>[</font><font color=#cc0000>"& oicq &"</font><font color=#000000>]</font><font color=green>介绍人</font><font color=#000000>[</font><font color=#cc0000>"& jsr &"</font><font color=#000000>]</font>(<font color=#ff00ff>" & regsj & "</font>)"			'聊天数据
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
Response.Write "<script Language=Javascript>alert('恭喜{" & name & "}，您已经成功加入本社区，请您登陆！\n以下信息请牢记，登陆后应及时申请密码保护！\n—— —— —— —— —— —— —— —— —— —— —— ——\n★ID：" & id & "\n姓名：" & name & "\n密码：" & pswc & "\n国家：" &country& "\n如果连续两个月未登陆\n您的ID将被新用户取代！');window.parent.close();</script>"
Response.End
%>