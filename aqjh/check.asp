<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<!--#include file="const2.asp"-->
<!--#include file="const3.asp"-->
<!--#include file="pass.asp"-->
<!--#include file="showchat.asp"-->
<!--#include file="chk.asp"-->
<%Response.Expires=0
call Chkproxy()
if aqjh_myie=1 then call Chkmyie()
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
call chkpost()
if Application("aqjh_room")="" then Response.Redirect "error.asp?id=000"
sername=Request.ServerVariables("SERVER_NAME")
ip=Request.ServerVariables("LOCAL_ADDR")
ip="127.0.0.1"
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
'if InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE")=0 then Response.Redirect "error.asp?id=010"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if aqjh_disproxy=1 then  Chkproxy()
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
if y<10 then
sjj=n & "-" & y+3 & "-" & r
else
sjj=n+1 & "-" & y+3-12 & "-" & r
end if
if y<11 then
sjjj=n & "-" & y+2 & "-" & r
else
sjjj=n+1 & "-" & y+2-12 & "-" & r
end if
huiqi=n+1 & "-" & y & "-" & r
userip=Request.ServerVariables("REMOTE_ADDR")
name=Trim(Request.Form("name"))
password=Trim(Request.Form("pass"))
if Application("aqjh_closedoor")="1" and name<>Application("aqjh_user") then Response.Redirect "error.asp?id=100"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
next
if onlinenow>Application("aqjh_chat_maxpeople") then Response.Redirect "error.asp?id=101"
if instr(aqjh_disloginname,name)<>0 then Response.Redirect "error.asp?id=130"
if session("aqjh_name")<>"" then
	Session.Abandon
	Response.End 
end if		
if session("aqjh_name")<>"" and session("aqjh_name")=name then 
	Response.Redirect "welcome.asp"
	Response.End 
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then 
		Session.Abandon
		Response.Redirect "error.asp?id=140"
		Response.End 
	end if
next 
if session("aqjh_name")<>"" then 
	Response.Redirect "welcome.asp"
	Response.End 
end if
if len(name)>5  then 
	Response.Write "<script language=JavaScript>{alert('提示：["&name&"]姓名长度，最多为10个字符！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if password="" or name="" then Response.Redirect "error.asp?id=128"
if server.URLEncode(password)<>password then Response.Redirect "error.asp?id=121"
if chuser(name) then Response.Redirect "error.asp?id=120"
badword="我塞,黑客,嘿客,老师,射精,奸,去死,老子,弟,吃屎,你妈,你娘,日你,鬼,卫情,江湖,中国,色,尻,操你,色狼,干死你,王八,逼,贱人,狗娘,婊子,表子,靠你,叉你,叉死,垃圾,混蛋,浑蛋,插你,插死,干你,干死,日死,鸡巴,睾丸,死去,处女,处男,按摸,爬你达来蛋,撅你达来蛋,死你达来蛋,包皮,龟头,猪,狗,猫,屄,赑,妣,肏,奶子,尻,屌,作爱,做爱,床上,抱抱,鸡八,打炮,十八摸,你爷,你奶奶,你老爸,你老妈,我做你,你爸,我儿,操你,逼,网管,掌门,我干,我操,你母,我搞,去死,法伦,法伦功,锦涛,腚,猪,踢,抓,爹,爸,我们,爷,父,妈,娘,奶,鞋,妓,娼,摸,阴,蒂,奸,死,屎,尻,操,逼,贱,狗,婊,表,靠,叉,插,干,龟,头,屄,赑,妣,肏,尻,屌,床,抱,鸡,蛋,炮,·,主席,哥,我靠,泽民,洪志,六扇门,官府,公安,警察,检察,法院,政府,司法,祖"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(name),bad(i))<>0 then Response.Redirect "error.asp?id=131"
next
dieip=aqjh_dieip
ipk=split(userip,".",-1)
if Instr(dieip,"*.*.*.*")<>0 or Instr(dieip,ipk(0)&".*.*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&".*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&"."&ipk(2)&".*")<>0 or Instr(dieip,userip)<>0 then Response.Redirect "error.asp?id=111"
iplocktime=int(Application("aqjh_iplocktime"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dcz=0
sql="SELECT a FROM i WHERE DateDiff('n',b,#" & sj & "#)>=" & iplocktime
rs.open sql,conn,1,1
if Not(rs.Eof and rs.Bof) then dcz=1
rs.close
if dcz=1 then
	conn.Execute "DELETE FROM i WHERE DateDiff('n',b,#" & sj & "#)>=" & iplocktime
end if
sql="SELECT a,b FROM i WHERE a='" & userip & "'"
rs.open sql,conn,1,1
if NOT(rs.Eof and rs.Bof) then
	lockdate=rs("b")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=110&lockdate=" & server.URLEncode(lockdate)
end if
rs.close
'转换密码：
password=md5(password)
sql="SELECT * FROM 用户 WHERE 姓名='"&name&"'"
rs.open sql,conn,2,2
if rs.Eof and rs.Bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=423"
	response.end
end if
if rs("密码")<>password then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=141"
	response.end
end if
aqjh_name=rs("姓名")
aqjh_grade=rs("grade")
aqjh_jhdj=rs("等级")
'reglastkick=rs("lastkick")
sjyy=rs("事件原因")
if rs("体力")<-100 or rs("状态")="死" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=421&sjyy="&sjyy
	response.end
end if
'对内力和武功超出下限的处理
if rs("内力")<-1000 then
	conn.execute("update 用户 set 内力=0 where 姓名='"&aqjh_name&"'")
end if
if rs("武功")<-1000 then
	conn.execute("update 用户 set 武功=0 where 姓名='"&aqjh_name&"'")
end if
if rs("攻击加")<0 then
	conn.execute("update 用户 set 攻击加=0 where 姓名='"&aqjh_name&"'")
end if
if rs("防御加")<0 then
	conn.execute("update 用户 set 防御加=0 where 姓名='"&aqjh_name&"'")
end if
'对通缉犯处理
if rs("道德")<-1000 or rs("魅力")<-1000 then
	conn.execute("update 用户 set 通缉=True,保护=false where 姓名='"&aqjh_name&"'")
end if
'对金钱超出上限的处理
mess1="<font color=red>"&aqjh_name&"</font>喝得醉醺醺的摇头晃脑刚想进江湖，"
if rs("银两")>1000000000 then
	mess="官府突然出现：你哪来这么多钱？[银两]都超过十亿征收个人所得税……"
        conn.execute "update 用户 set 银两=1000000000 where  姓名='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('提示：官府突然出现：你哪来这么多钱？[银两]都超过十亿征收个人所得税……');</script>"
	if aqjh_grade<10 then call showchat("<font color=ff0000>【官府收税】</font>" & mess1 & mess)
end if
if rs("存款")>1000000000 then
	mess="官府突然出现：你哪来这么多钱？[存款]都超过十亿征收个人所得税……"
        conn.execute "update 用户 set 存款=1000000000 where  姓名='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('提示：官府突然出现：你哪来这么多钱？[存款]都超过十亿征收个人所得税……');</script>"
	if aqjh_grade<10 then call showchat("<font color=ff0000>【官府收税】</font>" & mess1 & mess)
end if
if rs("金币")>1000000000 then
	mess="官府突然出现：你哪来这么多钱？[金币]都超过十亿征收个人所得税……"
        conn.execute "update 用户 set 金币=1000000000 where  姓名='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('提示：官府突然出现：你哪来这么多钱？[金币]都超过十亿征收个人所得税……');</script>"
	if aqjh_grade<10 then call showchat("<font color=ff0000>【官府收税】</font>" & mess1 & mess)
end if
'对金钱超出上限处理完毕
userid=rs("id")
if DateDiff("d",date(),rs("会员结束"))<=0 and rs("会员")=True then
	conn.Execute "update 用户 set 会员=False where 姓名='"&aqjh_name&"'"
	Response.Write "<script Language=Javascript>alert('提示：你的[泡点制会员]已经结束!');</script>"
end if
aqjh_hy=rs("会员等级")
aqjh_hydate=rs("会员日期")
hygq=DateDiff("d",date(),aqjh_hydate)
if aqjh_hy>0 and hygq<5 then
	Response.Write "<script Language=Javascript>alert('提示：你的会员还有"&hygq&"天，请你尽快与站长联系！');</script>"
end if
if aqjh_hy>0 and hygq<=0 then
	Response.Write "<script Language=Javascript>alert('提示：你的会员已经到期你的各项值请恢复到非会员状态！');</script>"
	hydj=aqjh_hy
	conn.Execute "update 用户 set 会员等级=0,职业='采冰' where 姓名='"&aqjh_name&"'"
end if
zt=trim(rs("状态"))
dldate=rs("登录")
denglucha=DateDiff("d",dldate,now())
if denglucha>10 and rs("grade")>5 and rs("grade")<9 then
conn.execute("update 用户 set grade=1,门派='游侠' where 姓名='" &aqjh_name& "'")
Response.Write "<script Language=javascript>alert('【快乐江湖】提示：你已经"& denglucha &"天未登录快乐江湖。\n6-8级官府超过10天未登录，自动撤职！');</script>"
end if
if zt="监禁" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=420&sjyy="&sjyy
	response.end
end if
if rs("登录")>now() and (zt="狱" or zt="牢" or zt="点穴" or zt="眠") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	select case zt
		case "狱"
				Response.Redirect "error.asp?id=420&sjyy="&sjyy
				response.end
		case "牢"
				Response.Redirect "error.asp?id=422&sjyy="&sjyy
				response.end
		case "点穴" 
				Response.Redirect "error.asp?id=480&sjyy="&sjyy
				response.end
		case "眠" 
				Response.Redirect "error.asp?id=472"
				response.end
	end select
end if
if rs("事件原因")<>"无" then
	Response.Write "<script Language=Javascript>alert('江湖公告:"&rs("事件原因")&"');</script>"
end if
aqjh_hy=rs("会员等级")
prevtime=CDate(rs("lasttime"))
value=clng(rs("allvalue"))
dengji=int(sqr(value/50))
ltime=rs("lasttime")
lip=rs("lastip")
if DateDiff("m",prevtime,sj)<>0 then
	sqlstr="update 用户 set mvalue=0,等级="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where 姓名='"&aqjh_name&"'"
else
	sqlstr="update 用户 set 等级="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where 姓名='"&aqjh_name&"'"
end if
conn.execute( sqlstr )
Session("aqjh_name")=rs("姓名")
Session("aqjh_grade")=rs("grade")
Session("aqjh_jhdj")=rs("等级")
Session("aqjh_id")=rs("id")
session("nowinroom")=0
Session.Timeout=60
response.cookies("aqjh").expires=date()+10 
response.cookies("aqjh")("aqjh_sid")=session.sessionid
if Session("aqjh_grade")>=10 and instr(Application("aqjh_admin"),Session("aqjh_name"))=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=478"
	response.end
end if
if day(dldate)<>day(now()) or month(dldate)<>month(now()) or year(dldate)<>year(now()) then
	conn.execute("update 用户 set 杀人数=0 where 姓名='" & aqjh_name & "'")
end if
conn.execute("update 用户 set 状态='正常',事件原因='无',登录=now() where 姓名='"&aqjh_name&"'")
hy_sx=0
if rs("会员")=True and DateDiff("h",prevtime,sj)<24 then
	hy_sx=1
end if
mess1="<font color=red>"&aqjh_name&"</font>在家里苦练了一个通宵的九阳真经，刚想进江湖，"
if aqjh_sx=1 and hy_sx<>1 then
	wgj=(dengji*aqjh_wgsx+3800)+rs("武功加")
        if rs("武功")>wgj then
        mess2="[武功]"
        messok=1
	conn.execute("update 用户 set 武功="& wgj &" where 姓名='"&aqjh_name&"'")
        end if
	nlj=(dengji*aqjh_nlsx+2000)+rs("内力加")
        if rs("内力")>nlj then
        mess3="[内力]"
        messok=1
	conn.execute("update 用户 set 内力="& nlj &" where 姓名='"&aqjh_name&"'")
        end if
	tlj=(dengji*aqjh_tlsx+5260)+rs("体力加")
        if rs("体力")>tlj then
	conn.execute("update 用户 set 体力="& tlj &" where 姓名='"&aqjh_name&"'")
        mess4="[体力]"
        messok=1
        end if
 if aqjh_grade<10 and messok=1 then
    call showchat("<font color=ff0000>【走火入魔】</font>"&mess1&"因为<font color=red>"&mess2&mess3&mess4&"</font>超出人体极限差点走火入魔……还好得到仙人的点化恢复正常")
 end if
'对攻击和防御处理
 mygj=(rs("等级")+1)*aqjh_gjsx+4500+rs("攻击加")
 myfy=(rs("等级")+1)*aqjh_fysx+3000+rs("防御加")
 if rs("攻击")>mygj then
	conn.execute "update 用户 set 攻击="& mygj &" where 姓名='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('提示：您的攻击已经超出上限!系统已经修正!');</script>"
 end if
 if rs("防御")>myfy then
	conn.execute "update 用户 set 防御="& myfy &" where  姓名='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('提示：您的防御已经超出上限!系统已经修正!');</script>"
 end if
end if
'计数器处理
conn.execute("update [count] set num=num+1 where count='计数器'")
   '=========================================================
   ' BbsXp 5.13b for 快乐江湖9.9插件
   '=========================================================
   dim bbsconn,rs1,rsbbs,UserQuesion,UserAnswer,passjh,readme,myboardid,jhtx
   db="bbs/database/aqjh_bbs.asp"
   Set bbsconn = Server.CreateObject("ADODB.Connection")
   Set rs1=Server.CreateObject("ADODB.RecordSet")
   Set rsbbs=Server.CreateObject("ADODB.RecordSet")
   connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
   bbsconn.open connstr
   sql="select * from [User] where username='"& session("aqjh_name") &"'"
   rsbbs.open sql,bbsconn,1,3
   if rsbbs.eof then
       rsbbs.addnew
       rsbbs("username")=rs("姓名")
       rsbbs("UserInfo")="\中国\\\\\\\\\\\\\"
       rsbbs("UserIM")="\\\\\"
       rsbbs("friend")="|"
   end if
   '修改身份，如果是官府掌门就设置为社区站长
   if rs("门派")="官府" and rs("身份")="掌门" and rs("姓名")=Application("aqjh_user") then
	Rsbbs("membercode")=5
	bbsconn.execute("update [clubconfig] set adminpassword='"&ucase(rs("密码"))&"'")
   end if
   rsbbs("honor")=rs("身份")
   rsbbs("faction")=rs("门派")
   rsbbs("usermail")=rs("信箱")
   if rs("性别")="女" then rsbbs("userface")="images/face/49.gif" else rsbbs("userface")="images/face/28.gif"
   randomize timer
   pass1=int(rnd*8998999)+1000
   pass2=int(rnd*8998999)+1000
   rsbbs("question")=mid(md5(pass1),1,10)
   rsbbs("answer")=mid(md5(pass2),1,10)
   if rs("性别")="女" then rsbbs("Sex")="female" else rsbbs("Sex")="male" end if
   Rsbbs("landtime")=NOW()
   rsbbs("userpass")=mid(rs("密码"),1,10)
   rsbbs.update
   rsbbs.close
   '取用户的数据写cookie为论坛用户用
        sql="select * from [User] where username='"&session("aqjh_name")&"'"
        rsbbs.open sql,bbsconn,1,1
        Response.Cookies("username")=rsbbs("username")
        Response.Cookies("userpass")=rsbbs("userpass")
	Response.Cookies("eremite")="0"
        rsbbs.close    
'论坛写入结束
'会员提示
rs.close
rs.open "SELECT * FROM sm where a='进入'",conn,2,2
banner=rs("c")
mygg=rs("d")
if aqjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('"&mygg&"\n\n您上次登陆的时间是："&ltime&"\n您上次登陆的IP是："&lip&"');</script>"
else
	banners=Split(Trim(banner),";",-1)
	total=UBound(banners)
	randomize timer
	x=int(rnd*(total+1))
	Response.Write "<script Language=Javascript>alert('"&banners(x)&"\n\n您上次登陆的时间是："&ltime&"\n您上次登陆的IP是："&lip&"');</script>"
	erase banners
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
if aqjh_myie=1 then
 call Chkmyie()
else
 Response.Write "<script Language=Javascript>chatwin=window.open('chat/jhchat.asp','aqjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);</script>"
 Response.Write "<script Language=Javascript>location.href = 'AQjh.HTM';</script>"
end if
%>