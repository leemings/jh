<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<!--#include file="pass.asp"-->
<%Response.Expires=0
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
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
'if mid(server_v1,8,len(server_v2))<>server_v2 then
'        response.write "严禁使用泡分工具！"
'        response.end
'end if
if Application("sjjh_room")="" then Response.Redirect "error.asp?id=000"
sername=Request.ServerVariables("SERVER_NAME")
ip=Request.ServerVariables("LOCAL_ADDR")
ip = "127.0.0.1"
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if sjjh_disproxy="1" and (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "error.asp?id=011"
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
userip = "127.0.0.1"
if Application("sjjh_closedoor")="1" then
Response.Write "<script Language=Javascript>alert('欢迎您的光临！\n由于最近服务器不太稳定，所以站长决定暂时关闭聊天室的登录功能\n现在不能进行登录，请过些时间再来！\n—— —— —— —— —— —— —— —— —— —— —— —— —— —— ——\n感谢您对本江湖的支持和厚爱！\n给您带来的不便，甚感抱歉！\nSEE YOU\n\n2003-11-10');window.parent.close();</script>"
Response.End
end if
name=Trim(Request.Form("name"))
password=Trim(Request.Form("pass"))
if instr(sjjh_disloginname,name)<>0 then Response.Redirect "error.asp?id=131"
if session("sjjh_name")<>"" then
	Session.Abandon
	Response.End 
end if
		
if session("sjjh_name")<>"" and session("sjjh_name")=name then 
	Response.Redirect "desktop/"
	Response.End 
end if
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("sjjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then 
		Session.Abandon
		Response.Redirect "error.asp?id=140"
		Response.End 
	end if
next 
if session("sjjh_name")<>"" then 
	Response.Redirect "desktop/"
	Response.End 
end if
if len(name)>5  then 
	Response.Write "<script language=JavaScript>{alert('提示：["&name&"]姓名长度，最多为10个字符！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if password="" or name="" then Response.Redirect "error.asp?id=128"
if server.URLEncode(password)<>password then Response.Redirect "error.asp?id=121"
if chuser(name) then Response.Redirect "error.asp?id=120"
badword="我塞,黑客,嘿客,老师,射精,奸,去死,老子,弟,吃屎,你妈,你娘,日你,鬼,一钱天,江湖,中国,色,尻,操你,色狼,干死你,王八,逼,贱人,狗娘,婊子,表子,靠你,叉你,叉死,垃圾,混蛋,浑蛋,插你,插死,干你,干死,日死,鸡巴,睾丸,死去,处女,处男,按摸,爬你达来蛋,撅你达来蛋,死你达来蛋,包皮,龟头,猪,狗,猫,屄,赑,妣,肏,奶子,尻,屌,作爱,做爱,床上,抱抱,鸡八,打炮,十八摸,你爷,你奶奶,你老爸,你老妈,我做你,你爸,我儿,操你,逼,网管,掌门,我干,我操,你母,我搞,去死,法伦,法伦功,锦涛,腚,猪,踢,抓,爹,爸,我,爷,父,妈,娘,奶,鞋,妓,娼,摸,阴,蒂,奸,死,屎,尻,操,逼,贱,狗,婊,表,靠,叉,插,干,龟,头,屄,赑,妣,肏,尻,屌,床,抱,鸡,蛋,炮,·,主席,哥,我靠,泽民,洪志,六扇门,官府,公安,警察,检察,法院,政府,司法,祖"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(name),bad(i))<>0 then Response.Redirect "error.asp?id=131"
next
dieip=sjjh_dieip
ipk=split(userip,".",-1)
if Instr(dieip,"*.*.*.*")<>0 or Instr(dieip,ipk(0)&".*.*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&".*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&"."&ipk(2)&".*")<>0 or Instr(dieip,userip)<>0 then Response.Redirect "error.asp?id=111"
iplocktime=int(Application("sjjh_iplocktime"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
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

sjjh_name=rs("姓名")
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

'对通缉犯处理
if rs("道德")<-1000 or rs("魅力")<-1000 then
	conn.execute("update 用户 set 通缉=True,保护=false where 姓名='"&sjjh_name&"'")
end if
if DateDiff("d",date(),rs("会员结束"))<=0 and rs("会员")=True then
	conn.Execute "update 用户 set 会员=False where 姓名='"&sjjh_name&"'"
	Response.Write "<script Language=Javascript>alert('提示：你的[泡点制会员]已经结束!');</script>"
end if

sjjh_jhdj=rs("等级")
mygj=rs("等级")*sjjh_gjsx+4500 
myfy=rs("等级")*sjjh_fysx+3000
if rs("攻击")>mygj then
	conn.execute "update 用户 set 攻击="& mygj &" where 姓名='"&sjjh_name&"'"
end if
if rs("防御")>myfy then
	conn.execute "update 用户 set 防御="& myfy &" where  姓名='"&sjjh_name&"'"
end if

if sjjh_jhdj=60 and rs("会员等级")=1 then
conn.Execute "update 用户 set 会员等级=2,会员日期='"&sjj&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你等级达60，是2级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=90 and rs("会员等级")=2 then
conn.Execute "update 用户 set 会员等级=3,会员日期='"&sjj&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你等级达90，是3级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=90 and rs("会员等级")=1 then
conn.Execute "update 用户 set 会员等级=2,会员日期='"&sjj&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你，是2级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=120 and rs("会员等级")=3 then
conn.Execute "update 用户 set 会员等级=4,会员日期='"&huiqi&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你等级达120，是4级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=120 and rs("会员等级")=2 then
conn.Execute "update 用户 set 会员等级=3,会员日期='"&huiqi&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你，是3级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=120 and rs("会员等级")=1 then
conn.Execute "update 用户 set 会员等级=2,会员日期='"&huiqi&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你，是2级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("会员等级")=4 then
conn.Execute "update 用户 set 会员等级=5,会员日期='"&huiqi&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你等级达300，是5级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("会员等级")=3 then
conn.Execute "update 用户 set 会员等级=4,会员日期='"&huiqi&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你，是4级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("会员等级")=2 then
conn.Execute "update 用户 set 会员等级=3,会员日期='"&huiqi&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你，是3级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("会员等级")=1 then
conn.Execute "update 用户 set 会员等级=2,会员日期='"&huiqi&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：恭喜你，是2级会员了！记得多介绍朋友来玩呀 ^_^');</script>"
end if
if rs("会员等级")=0 and rs("会员")=false then
conn.Execute "update 用户 set 会员等级=1,会员日期='"&sjjj&"' where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('提示：站长再次送您1级会员两个月！要努力了 ^_^');</script>"
end if

yin=rs("银两")
cun=rs("存款")
if (yin+cun)>2000000000 then
conn.Execute "update 用户 set 银两=100000000,存款=900000000 where 姓名='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('系统提醒：您的银两与存款总和超过20亿！系统自动为您清为10亿！以后要注意了！ ^_^');</script>"
end if

sjjh_hy=rs("会员等级")
sjjh_hydate=rs("会员日期")
hygq=DateDiff("d",date(),sjjh_hydate)
if sjjh_hy>0 and hygq<5 then
	Response.Write "<script Language=Javascript>alert('提示：你的会员还有"&hygq&"天，请你尽快与站长联系！');</script>"
end if
if sjjh_hy>0 and hygq<=0 then
	Response.Write "<script Language=Javascript>alert('提示：你的会员已经到期你的各项值请恢复到非会员状态！');</script>"
	hydj=sjjh_hy
	Select Case hydj
	case 1
		conn.Execute "update 用户 set 会员等级=0 where 姓名='"&sjjh_name&"'"
	case 2
		conn.Execute "update 用户 set 会员等级=0 where 姓名='"&sjjh_name&"'"
	case 3
		conn.Execute "update 用户 set 会员等级=0 where 姓名='"&sjjh_name&"'"
	case 4
		conn.Execute "update 用户 set 会员等级=0,grade=1,门派='游侠',身份='弟子' where 姓名='"&sjjh_name&"'"
	end select
end if
zt=trim(rs("状态"))
dldate=rs("登录")
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
sjjh_hy=rs("会员等级")
prevtime=CDate(rs("lasttime"))
value=clng(rs("allvalue"))
dengji=int(sqr(value/50))
if DateDiff("m",prevtime,sj)<>0 then
	sqlstr="update 用户 set mvalue=0,等级="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where 姓名='"&sjjh_name&"'"
else
	sqlstr="update 用户 set 等级="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where 姓名='"&sjjh_name&"'"
end if
conn.execute( sqlstr )
Session("sjjh_id")=rs("id")
Session("sjjh_name")=rs("姓名")
Session("sjjh_grade")=rs("grade")
Session("sjjh_jhdj")=rs("等级")
session("nowinroom")=0
if rs("姓名")=Application("sjjh_user") then
session("b28")="ok"
end if
Session.Timeout=300
response.cookies("yxjh").expires=date()+10 
response.cookies("yxjh")("sjjh_sid")=session.sessionid
if Session("sjjh_grade")>=10 and instr(Application("sjjh_admin"),Session("sjjh_name"))=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=478"
	response.end
end if
if day(dldate)<>day(now()) or month(dldate)<>month(now()) or year(dldate)<>year(now()) then
	conn.execute("update 用户 set 杀人数=0 where 姓名='" & sjjh_name & "'")
end if
conn.execute("update 用户 set 状态='正常',事件原因='无',登录=now() where 姓名='"&sjjh_name&"'")
hy_sx=0
if rs("会员")=True and DateDiff("h",prevtime,sj)<24 then
	hy_sx=1
end if
if sjjh_sx=1 and hy_sx<>1 then
	wgj=(dengji*sjjh_wgsx+3800)+rs("武功加")
	conn.execute("update 用户 set 武功="& wgj &" where 武功>"& wgj &" and 姓名='"&sjjh_name&"'")
	nlj=(dengji*sjjh_nlsx+2000)+rs("内力加")
	conn.execute("update 用户 set 内力="& nlj &" where 内力>"& nlj &" and 姓名='"&sjjh_name&"'")
	tlj=(dengji*sjjh_tlsx+5260)+rs("体力加")
	conn.execute("update 用户 set 体力="& tlj &" where 体力>"& tlj &" and 姓名='"&sjjh_name&"'")
end if
'计数器处理
conn.execute("update sm set b=b+1 where a='计数器'")

'先检查用户是否登录江湖及在论坛中是否有江湖用户的资料，如果没有就加入
   dim bbsconn,rs1,rsbbs,question,answer,passjh,readme,myboardid
   db="bbs/data/Eline_bbs_6.3.0.asp"
   Set bbsconn = Server.CreateObject("ADODB.Connection")
   Set rs1=Server.CreateObject("ADODB.RecordSet")
   Set rsbbs=Server.CreateObject("ADODB.RecordSet")
   connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
   '如果你的服务器采用较老版本Access驱动，请用下面连接方法
   'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)
   bbsconn.open connstr

   sql="select * from [user] where username='"&session("sjjh_name") &"'"
   rsbbs.open sql,bbsconn,1,3
   if rsbbs.eof then
       rsbbs.addnew
       rsbbs("Article")=0
       set rs1=bbsconn.execute("select usertitle,titlepic from usertitle where not minarticle=-1 order by minarticle")
       rsbbs("userclass")=rs1(0)
       rsbbs("titlepic")=rs1(1)
       rs1.close
  
       rsbbs("showRe")=0
       rsbbs("logins")=1

	randomize timer
   	passjh=int(rnd*8998999)+1000
   	question=int(rnd*8998999)+1000
   	answer=int(rnd*8998999)+1000

   	rsbbs("username")=rs("姓名")

   	rsbbs("userpassword")=mid(md5(passjh),1,10)
        rsbbs("userWealth")=600
   	rsbbs("useremail")=rs("信箱")
   	rsbbs("oicq")=rs("oicq")
   	rsbbs("quesion")=mid(md5(question),1,10)
   	rsbbs("answer")=mid(md5(answer),1,10)
   end if
   '修改身份，如果是官府掌门就设置为1为master，如果是其它的掌门就设置了3 boardmaster
   rsbbs("usergroupid")=4
   if rs("门派")="官府" and rs("身份")="掌门" then  Rsbbs("usergroupid")=1
   if rs("门派")<>"官府" and rs("身份")="掌门" then  Rsbbs("usergroupid")=3
   rsbbs("title")=rs("身份")
   rsbbs("usergroup")=rs("门派")
   if rs("性别")="女"  then rsbbs("sex")=0 else rsbbs("sex")=1 end if
   Rsbbs("lastlogin")=NOW()
   rsbbs.update
   rsbbs.close
   '再检查官府掌门是不admin表中，如果没有就加入

   if rs("门派")="官府" and rs("身份")="掌门" then
	sql="select * from admin where username='"&session("sjjh_name")&"'"
   	rsbbs.open sql,bbsconn,1,3
   	if rsbbs.eof then
   	   rsbbs.addnew
     	   rsbbs("username")=session("sjjh_name")
   	end if
     	rsbbs("password")=rs("密码")
     	rsbbs("flag")="01, 02, 03, 04, 11, 12, 13, 14, 15, 16, 17, 21, 22, 23, 24, 25, 26, 27, 28, 31, 32, 33, 34, 35, 41, 42, 43, 44, 51, 52, 53, 54, 55, 61, 62, 63, 64, 71, 72, 73, 74, 75"
   	rsbbs("adduser")=session("sjjh_name")
   	rsbbs.update
   	rsbbs.close
   end if
   '取用户的数据写cookie为论坛用户用
        dim cookies_path_s,cookies_path_d,cookies_path,usercookies
   	'判断更新cookies目录
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	for i=1 to cookies_path_d-1
		if not (cookies_path_s(i)="upload" or cookies_path_s(i)="admin") then 	cookies_path=cookies_path&cookies_path_s(i)&"/"
		
	next
	if cookiepath<>cookies_path then
		cookiepath=cookies_path
		bbsconn.execute("update config set cookiepath='"&cookiepath&"'")
	end if

        sql="select userid,userpassword,userclass from [user] where username='"&session("sjjh_name")&"'"
        rsbbs.open sql,bbsconn,1,1
	Response.Cookies("aspsky")("username") = session("sjjh_name")
	Response.Cookies("aspsky")("userid") = rsbbs("userid")
	Response.Cookies("aspsky")("password") = rsbbs("userpassword")
	Response.Cookies("aspsky")("userclass") = rsbbs("userclass")
	Response.Cookies("aspsky")("userhidden") = 2
	Response.Cookies("aspsky")("usercookies")="0"
	rem 清除图片上传数的限制
	response.cookies("upNum")=0
	Response.Cookies("aspsky").path=cookiepath
	
	rsbbs.close
   '检查门派是否在groupname表中，如果没有就加入
   if rs("门派")<>"官府" and rs("身份")="掌门" then
     sql="select * from groupname where groupname='"&rs("门派")&"'"
     rsbbs.open sql,bbsconn,1,3
     if rsbbs.eof then
       rsbbs.addnew
       rsbbs("groupname")=rs("门派")
       rsbbs.update
     end if
     rsbbs.close
     '查找这个门派是否有自己的论坛？
     sql="select * from board where boardtype='『 "&rs("门派")&" 』'"
     rsbbs.open sql,bbsconn,1,3
     if rsbbs.eof then
       '取七门八派
       '取门派的描述
       set rs1=conn.execute("select d from p where a='"&rs("门派")&"'")
       readme=rs1(0)
       rs1.close
       set rs1=bbsconn.execute("select max(boardid)+1 from board")
       myboardid=rs1(0)
       rs1.close
       set rs1=bbsconn.execute("select boardid,rootid,orders from board where boardtype='七门八派'")
       if not rs1.eof then
          rsbbs.addnew
          rsbbs("boardid")=myboardid
          rsbbs("boardtype")="『 "&rs("门派")&" 』"
          rsbbs("parentid")=rs1(0)
          rsbbs("parentstr")=rs1(0)
          rsbbs("depth")=1
          rsbbs("rootid")=rs1(1)
          rsbbs("child")=0
          rsbbs("orders")=rs1("orders")+1
          rsbbs("readme")=readme
          rsbbs("boardmaster")=""
          rsbbs("lastbbsnum")=0
          rsbbs("lasttopicnum")=0
          rsbbs("indeximg")=""
          rsbbs("todaynum")=0
          rsbbs("boarduser")=""
          rsbbs("lastpost")="$0$"&date()&" "&now()&"$$$$$"
          rsbbs("sid")=1
          rsbbs("board_setting")="0,0,0,0,1,0,1,1,1,1,1,1,1,1,1,1,16240,3,300,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,0|24,1,0,300,20,10,9,12,1,10,10,0,0,0,0,1,5,0,1,4,0,0,0,0,0,0,0,0,0"
          rsbbs.update
       end if
       rs1.close
       
       rsbbs.close
     end if
     
   end if
   set rsbbs=nothing
   set rs1=nothing

'会员提示
rs.close
rs.open "SELECT * FROM sm where a='进入'",conn,2,2
banner=rs("c")
mygg=rs("d")
if sjjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('"&mygg&"');</script>"
else
	banners=Split(Trim(banner),";",-1)
	total=UBound(banners)
	randomize timer
	x=int(rnd*(total+1))
	Response.Write "<script Language=Javascript>alert('"&banners(x)&"');</script>"
	erase banners
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>chatwin=window.open('chat/jhchat.asp','sjjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);</script>"
if Request.form("txwz")=0 then
Response.Write "<script Language=Javascript>location.href = 'skin.asp';</script>"
end if
if Request.form("txwz")=1 then
Response.Write "<script Language=Javascript>location.href = 'mh.asp';</script>"
end if
if Request.form("txwz")=2 then
Response.Write "<script Language=Javascript>location.href = 'cww.asp';</script>"
end if
if Request.form("txwz")=3 then
Response.Write "<script Language=Javascript>location.href = 'gs.asp';</script>"
end if
if Request.form("txwz")=4 then
Response.Write "<script Language=Javascript>location.href = 'szwd.asp';</script>"
end if
if Request.form("txwz")=5 then
Response.Write "<script Language=Javascript>location.href = 'sm.asp';</script>"
end if
if Request.form("txwz")=6 then
Response.Write "<script Language=Javascript>location.href = 'lr.asp';</script>"
end if
if Request.form("txwz")=7 then
Response.Write "<script Language=Javascript>location.href = 'qs.asp';</script>"
end if
if Request.form("txwz")=8 then
Response.Write "<script Language=Javascript>location.href = 'th.asp';</script>"
end if
if Request.form("txwz")=9 then
Response.Write "<script Language=Javascript>location.href = 'ppb.asp';</script>"
end if
if Request.form("txwz")=10 then
Response.Write "<script Language=Javascript>location.href = 'neweline.asp';</script>"
end if
if Request.form("txwz")=11 then
Response.Write "<script Language=Javascript>location.href = 'desktop/';</script>"
else
Response.Write "<script Language=Javascript>location.href = 'mh.asp';</script>"
end if
%>