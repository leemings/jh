<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'学习
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
response.buffer=false
response.write " "
tfjh_name=session("tfjh_name")
if Instr(LCase(" " & application("tfjh_useronlinename"&nowinroom)),LCase(" "&tfjh_name&" "))=0 then 
	session("tfjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if tfjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says," ")
fnn1=mid(says,i+1)
fnn1=trim(fnn1)
set rs=session("bug_conn_npc").execute("select 招式,魔法,w9,等级,结婚次数 from 用户 where 姓名='" & tfjh_name & "'")
if instr(";" & rs("w9"),";" & fnn1 & "|")=0 then response.write "<script>alert('你没有练习[" & fnn1 & "]的书啊！');</script>" :  response.end
select case fnn1				'判断是不是魔法
case "治愈术","心灵启示","定身术","天眼开光","圣言术","诱惑之光","集体疗伤","地震术","吸灵咒","北冥神功","浴灵","圣浴灵","浴血","圣浴血","净身"
if instr("," & rs("魔法"),"," & fnn1 & "|")>0 then response.write "<script>alert('你已经学会这种魔法了，没有必要再学了');</script>" :  response.end
select case fnn1
	case "治愈术"
		if rs("等级")<30 then response.write "<script>alert('你的战斗等级小于30级，无法修炼如此之高深的魔法~！');</script>" :  response.end
		mess1="##经过一番苦心修炼，终于学会了治愈术<img src='../shenbing/mofa/zhiyu.gif'>"
	case "心灵启示"
		if rs("等级")<20 then response.write "<script>alert('你的战斗等级小于20级，无法修炼如此之高深的魔法~！');</script>" :  response.end
		mess1="##刻苦地钻研了很久，终于领悟到了心灵启示<img src='../shenbing/mofa/xinling.gif'>的要领"
	case "定身术"
		if rs("等级")<100 then response.write "<script>alert('你的战斗等级小于100级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##仔细地把定身术的课本看了一遍又一遍，终于基本上搞懂了定身术<img src='../shenbing/mofa/dingshen.gif'>的用法"
	case "天眼开光"
		if rs("等级")<90 then response.write "<script>alert('你的战斗等级小于90级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##经过无数次失败，终于在无数+1次的时候学会了天眼开光<img src='../shenbing/mofa/haidi.gif' width=80>"
	case "圣言术"
		if rs("等级")<150 then response.write "<script>alert('你的战斗等级小于150级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##把嗓子都喊哑了，吃了好几片金嗓子喉宝，终于能喊出圣言了<img src='../shenbing/mofa/shengyan.gif' width=80>，学会了圣言术"
	case "诱惑之光"
		if rs("等级")<100 then response.write "<script>alert('你的战斗等级小于100级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##以惊人的毅力苦苦练习了N年，终于练成了诱惑之光<img src='../shenbing/mofa/youhuo.gif' width=80>"
	case "集体疗伤"
		if rs("等级")<250 then response.write "<script>alert('你的战斗等级小于250级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##研究了半天集体疗伤术，一个偶然机会，使##一下次练成了集体疗伤<img src='../shenbing/mofa/jiti.gif' width=80>"
	case "地震术"
		mess1="##把课本看了一遍，轻松学会了地震术！<img src='../shenbing/mofa/dizhen.gif' width=80>"
	case "吸灵咒"
		if rs("等级")<20 then response.write "<script>alert('你的战斗等级小于20级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##把课本看了一遍，轻松学会了吸灵咒！<img src='../shenbing/mofa/xiling.gif' width=80>"
	case "北冥神功"
		if rs("等级")<50 then response.write "<script>alert('你的战斗等级小于50级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##把课本看了一遍，轻松学会了北冥神功！<img src='../shenbing/mofa/beiming.gif' width=80>"
	case "浴灵"
		if rs("结婚次数")<>0 then response.write "<script>alert('你已经不是处子之身，不能修炼这种魔法了！');</script>" :  response.end
		mess1="##拿着课本跑到东大街的澡堂，洗了洗脚，就学会了浴灵！"
	case "浴血"
		if rs("结婚次数")<>0 then response.write "<script>alert('你已经不是处子之身，不能修炼这种魔法了！');</script>" :  response.end
		mess1="##拿着课本跑到西大街的澡堂，洗了洗手，就学会了浴血！"
	case "圣浴灵"
		if instr(rs("魔法"),"浴灵")=0 then response.write "<script>alert('你必须先学浴灵！');</script>" :  response.end
		if rs("结婚次数")=0 then response.write "<script>alert('你还是处子之身，不能修炼这种魔法！');</script>" :  response.end
		mess1="##拿着课本跑到南大街的澡堂，洗了洗脚，就学会了圣浴灵！"		
	case "圣浴血"
		if instr(rs("魔法"),"浴血")=0 then response.write "<script>alert('你必须先学浴血！');</script>" :  response.end
		if rs("结婚次数")=0 then response.write "<script>alert('你还是处子之身，不能修炼这种魔法！');</script>" :  response.end
		mess1="##拿着课本跑到北大街的澡堂，洗了洗手，就学会了圣浴血！"
	case "净身"
		if rs("等级")<50 then response.write "<script>alert('你的战斗等级小于50级，无法修炼如此之高深的魔法~！');</script>" :  response.end 
		mess1="##自己研究了课本，回到家里卫生间刷了刷牙，轻松悟出了[净身]！"
end select
tmp=trim(cstr(rs("魔法") & " "))
if tmp="" then 
	tmp=fnn1 & "|1"
else
	tmp=tmp & "," & fnn1 & "|1"
end if
session("bug_conn_npc").execute("update 用户 set 魔法='" & tmp & "',w9='" & abate(rs("w9"),fnn1,1) & "' where 姓名='" & tfjh_name & "'")
case else					'学习武功招式
set rsy=session("bug_conn_npc").execute("select * from y where b='NPC' and a='" & fnn1 & "'")
if rsy.eof then response.write "<script>alert('没有这种武功！数据有点不对！');</script>" :  response.end
if instr("," & rs("招式"),"," & fnn1 & "|")>0 then response.write "<script>alert('你已经学会这种武功了，没有必要再学了');</script>" :  response.end
tmp=trim(cstr(rs("招式") & " "))
if tmp="" then 
	tmp=fnn1 & "|1"
else
	tmp=tmp & "," & fnn1 & "|1"
end if
session("bug_conn_npc").execute("update 用户 set 招式='" & tmp & "',w9='" & abate(rs("w9"),fnn1,1) & "' where 姓名='" & tfjh_name & "'")
mess1="##看了看课本，经过不懈的努力终于练成了[" & fnn1 & "]"
end select
says="<font color=red>【学习】</font><b>" & mess1 & "</b>"  
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
%>
