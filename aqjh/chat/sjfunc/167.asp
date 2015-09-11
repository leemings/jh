<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'抢劫金库
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatroomsn=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
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
says="<font color=red>【抢劫金库】<font color=" & saycolor & ">"+qqyh()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function qqyh()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,通缉,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
yinl=rs("金币")
randomize timer
qjyl=int(rnd*yinl/100)
if rs("通缉")=True then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="[" & aqjh_name & "]" &":通缉犯还敢来抢金库，也太大胆了吧，小心抓你!"                                            	exit function
end if
if aqjh_jhdj<800 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="[" & aqjh_name & "]" &":你还是江湖小辈，就凭你这水平来抢金库，不是找死？还是等800级后再来!!"
	exit function
end if 
if  aqjh_grade>6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="[" & aqjh_name & "]" &":你是有身份地位的人，怎么能为了这么点小利去冒险!!只有管理6级以下的人才可以去抢!"
	exit function
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<=3000 then
	ss=3000-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="["& aqjh_name & "]:官府正在捉拿抢劫犯呢，你现在又去抢，想找死？还是再过"& ss &"秒钟再来!"
	exit function
end if
randomize timer
r=Int(Rnd*7)
select case r
	case 0
		yinl=yinl+qjyl
		qqyh="["& aqjh_name & "]从身后拽出一把<img src='../chat/PIC/273.gif'>,冲进中央银行金库，金库保镖一看，自己的无情剑<img src='../hcjs/jhjs/images/14.jpg'>也太落后了，吓得就跑，都忘了报警；结果[" & aqjh_name &"]抢得金币:"& qjyl
	case 1
		qqyh="["& aqjh_name & "]从身后拽出一把无情剑<img src='../hcjs/jhjs/images/14.jpg'>冲进中央银行金库，银行金库保镖一把拖出<img src='../chat/PIC/273.gif'>；嘿嘿，你也太落后了吧，[" & aqjh_name &"]只得束手就擒，把身上金币"&yinl&"全送给保镖，免了牢狱之灾。"
		yinl=0
	case 2
		yinl=(yinl+int(qjyl/2))
		qqyh="["& aqjh_name & "]从身后拽出一把无情剑<img src='../hcjs/jhjs/images/14.jpg'>冲进中央银行金库，在内应银行金库保镖的接应下[" & aqjh_name &"]抢得金币:" & qjyl & "分一半给了银行保镖"
	case 3
		yinl=yinl+500
		qqyh="["& aqjh_name & "]到了银行金库后发现无从下手，不过却发现天上掉下个林妹妹，意外捡到别人掉了的金币500个，高兴坏了！"
	case 4
		yinl=0
		qqyh="["& aqjh_name & "]从身后拽出一把<img src='../chat/pic/273.gif'>,冲进中央银行金库，银行金库保镖一看，自己的无情剑<img src='../hcjs/jhjs/images/14.jpg'>也太落后了，吓得就跑，一脚却踩到报警器；结果"& "[" & aqjh_name &"]抢得金币:" & qjyl & ",不幸的是，外面已经被官府人员包围了，幸好还认得官府人员，把身上的金币上交后得以走人"
	case 5
		yinl=int(yinl/2)
		qqyh="["& aqjh_name & "]到了银行金库后发现无从下手，不过却发现天上掉下个林妹妹，拿出自己身上的一半钱给金库支行长，却糊里糊涂地走出银行，支行行长不知道是什么意思，怕是别人来贿赂的就把金币上交给总行"
	case else
		qqyh="["& aqjh_name & "]:现在金库已经下班了，金币早被送到金库里去了，一分钱也没有，你来抢个什么呀。"
end select
conn.execute "update 用户 set 金币="&yinl&",操作时间=now() where 姓名='" & aqjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>