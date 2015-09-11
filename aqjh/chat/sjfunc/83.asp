<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'拐骗少女
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('提示：当女贩的等级最少得20级才行！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="大家" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>【拐骗少女】</font><font color=" & saycolor & ">"+fmrk(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'拐骗少女
function fmrk(to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select grade,银两,职业,身份,门派 from 用户 where 姓名='" & aqjh_name &"'",conn
if rs("职业")<>"女贩" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的职业不是女贩，不能操作！');}</script>"
	Response.End
end if
if rs("银两")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有带着一百万两银子！小心被抓住后一辈子当牛当马都还不清！');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你身为官府还想贩卖人口？你不想作了？');}</script>"
	Response.End
end if
if rs("门派")="出家" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你身为出家人还想做这种道德败坏的事情，太不道德了？');}</script>"
	Response.End
end if
rs.close
rs.open "select ID,grade,银两,性别,配偶,门派,魅力 from 用户 where 姓名='"& to1 &"'",conn
if rs("性别")<>"女" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不是看错了吧，男的你也想往红园里骗？');}</script>"
	Response.End
end if
if rs("门派")="出家" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：晕~~~~~你连出家人也想骗吗？');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不想活了，官府的你也敢骗！');}</script>"
	Response.End
end if
toid=rs("id")
peiou=rs("配偶")
if peiou=aqjh_name then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不会是连自已的老婆也想卖吧！');}</script>"
	Response.End
end if
%>
<!--#include file="data.asp"-->
<%
sql="select * from 妓女 where ID="&toid
set rs1=connt.execute(sql)
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	connt.close
	set connt=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你眼花了吧，她早就是红园的姑娘了！');}</script>"
	Response.End
end if
rs1.close
randomize timer
r=int(rnd*10)+1
fn1="拐骗少女"
meimao=rs("魅力")
yinliang=int(rs("银两")/2)
select case r
	case 1
		conn.execute "update 用户 set 银两=银两-1000000,道德=道德-500,魅力=魅力-500 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 银两=银两+1000000 where 姓名='" & to1 & "'"
		fmrk=aqjh_name & "想骗人没成功，反被"& to1 &"臭骂一通，急忙付了100万两银子赔情道歉，灰溜溜地跑了，道德下降500，魅力下降500。"
	case 2
		conn.execute "update 用户 set 银两=银两+500000,道德=道德-1000 where 姓名='" & aqjh_name & "'" 
		sql="insert into 妓女(妓女ID,姓名,美貌度,介绍人) values ("& toid &",'" & to1 & "'," & meimao &",'"&aqjh_name&"')"
set rs1=connt.execute(sql)
		fmrk=aqjh_name & "的嘴皮子真是好用，连蒙带骗的就把"& to1 &"卖到红院了，得到好处费50万，道德下降1000。" 
	case 3
		conn.execute "update 用户 set 状态='狱',登录=now()+1/144,事件原因='拐骗人口不成功被抓' where 姓名='" & aqjh_name & "'"
		conn.execute "insert into l(b,c,d,a,e) values ('系统','" & to1 & "','坐牢',now(),'" & fn1 & "')"
		fmrk=aqjh_name & "正在蒙骗"& to1 &"，却被官府人员发现，抓去坐牢了。10分钟后才可以登录。"
		call boot(aqjh_name,"坐牢，操作者："&aqjh_name&","&fn1)
	case 4
		yingliang=yingliang/2
		conn.execute "update 用户 set 银两=银两+" &yinliang& ",道德=道德-500,魅力=魅力-200 where 姓名='"& aqjh_name &"'"
		conn.execute "update 用户 set 银两="& yinliang &" where 姓名='" & to1 & "'"
		fmrk=aqjh_name & "想把"& to1 &"骗到红园没成功，但却骗到她所有现金的一半"& yinliang &","& aqjh_name &"道德下降500点，魅力下降200点。"
	case 5
		conn.execute "update 用户 set 银两=银两-1000000,道德=道德-500,魅力=魅力-500 where 姓名='"&aqjh_name&"'"
		fmrk=aqjh_name & "想诈骗"& to1 &"，谁知"& to1 &"和官府人员认识，"& aqjh_name &"赶快上下打典了100万两银子才算逃脱，道德和魅力因此各下降500点。"
	case 6
		conn.execute "update 用户 set 银两=银两+500000,道德=道德-1000 where 姓名='"&aqjh_name&"'" 
		sql="insert into 妓女(妓女ID,姓名,美貌度,介绍人) values ("& toid &",'" & to1 & "'," & meimao&",'"&aqjh_name&"')"
set rs1=connt.execute(sql)
		fmrk=aqjh_name & "的嘴皮子真是好用，连蒙带骗的就把"& to1 &"卖到红院了，得到好处费50万，道德下降1000。" 
	case 7
		conn.execute "update 用户 set 银两=0,道德=道德-500,体力=体力-30000,内力=内力-30000 where 姓名='"&aqjh_name&"'"
		fmrk=aqjh_name & "刚刚将"& to1 &"蒙骗住，却被"& to1 &"的一个仰慕者暴打一顿，损失所有现金，道德下降500，体力下降30000，内力下降30000。"
     case 8
                yingliang=yingliang/2
		conn.execute "update 用户 set 银两=银两+" &yinliang& ",道德=道德-500,魅力=魅力-200 where 姓名='"& aqjh_name &"'"
		conn.execute "update 用户 set 银两="& yinliang &" where 姓名='" & to1 & "'"
		fmrk=aqjh_name & "想把"& to1 &"骗到丽春青楼没成功，但却骗到她所有现金的一半"& yinliang &","& aqjh_name &"道德下降500点，魅力下降200点。"
     case 9
		conn.execute "update 用户 set 状态='狱',登录=now()+1/144,事件原因='拐骗人口不成功被抓' where 姓名='" & aqjh_name & "'"
		conn.execute "insert into l(b,c,d,a,e) values ('系统','" & to1 & "','坐牢',now(),'" & fn1 & "')"
		fmrk=aqjh_name & "正在蒙骗"& to1 &"，却被官府人员发现，抓去坐牢了。10分钟后才可以登录。"
		call boot(aqjh_name,"坐牢，操作者："&aqjh_name&","&fn1)
     case 10
                conn.execute "update 用户 set 银两=0,道德=道德-500,体力=体力-30000,内力=内力-30000 where 姓名='"&aqjh_name&"'"
		fmrk=aqjh_name & "刚刚将"& to1 &"蒙骗住，却被"& to1 &"的一个仰慕者暴打一顿，损失所有现金，道德下降500，体力下降30000，内力下降30000。"
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
connt.close
set connt=nothing
end function
%>