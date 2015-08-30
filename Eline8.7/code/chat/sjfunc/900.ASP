<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'振臂一呼♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
 Session("sjjh_inthechat")="0" 
 Response.Redirect "chaterr.asp?id=001" 
end if
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<21) and fjname="高手房间" then
 Response.Write "<script language=javascript>{alert('提示：现在是夺宝时间，不可以使用振臂一呼！');}</script>"
 Response.End
end if
if sjjh_jhdj<60 or sjjh_grade<2 then
 Response.Write "<script language=javascript>{alert('提示：振臂一呼需要60级，管理等级2级以上才可以操作！');}</script>"
 Response.End
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
 says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<bgsound src=wav/zb.wav loop=1><img src=img/zb.gif><font color=red><b>〖振臂一呼〗</b><font color=" & saycolor & ">"+zbyh()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'振臂一呼
function zbyh()
nowinroom=session("nowinroom")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select grade,等级,门派,身份,保留1,名单头像 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
bliu=rs("保留1")
grade=rs("grade")
mp=rs("门派")
sf=rs("身份")
touxiang=rs("名单头像")
rs.close
select case sf
 case "掌门"
  wgxs=0.02
  gjxs=0.02
  fyxs=0.02
 case "长老"
  wgxs=0.015
  gjxs=0.015
  fyxs=0.015
 case "护法"
  wgxs=0.01
  gjxs=0.01
  fyxs=0.01
 case "堂主"
  wgxs=0.005
  gjxs=0.005
  fyxs=0.005
end select
if mp="官府" or mp="天网" or mp="出家" or mp="游侠" or grade<2 or (sf<>"掌门" and sf<>"护法" and sf<>"堂主" and sf<>"长老") then
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('天网、官府、游侠、出家人不可以使用振臂一呼功能，并且要求身份为掌门、长老、护法或者堂主才可以使用！');}</script>"
 Response.End
end if
zb3=0 '加攻击
zb4=0 '加防御
zb5=0 '加武功
if bliu<>"保留" then
 zbdata=split(bliu,"|")
 zs=ubound(zbdata)
 if zs=4 then
  zb1=clng(zbdata(0)) '使用次数
  zb2=zbdata(1) '最后一次使用时间
  erase zbdata
  sj=DateDiff("d",zb2,now())

  if zb1>=3 and sj<0 then
  zbz=0
  zbb=now()
  zb3=0 '加攻击
  zb4=0 '加防御
  zb5=0 '加武功
  bliuu=zbz&"|"&zbb&"|"&zb3&"|"&zb4&"|"&zb5
  conn.execute "update 用户 set 保留1='"& bliuu &"' where 姓名='"&sjjh_name&"'"
  Response.Write "<script Language=Javascript>alert('提示：数据清零，你一天可以使用三次振臂一呼！');</script>"
  Response.End
  end if

  if zb1>=3 and sj=0 then
   set rs=nothing
   conn.close
   set conn=nothing
   Response.Write "<script language=javascript>{alert('你今天已经使用过三次振臂一呼功能了，每天最多可以使用三次！');}</script>"
   Response.End
  end if
 else
  zb1=0 '使用次数
  zb2=now() '最后一次使用时间
 end if
end if
'遍历在线弟子数
rs.open "select 姓名,武功,攻击,防御 from 用户 where 门派='"&mp&"' and grade<"& grade
if rs.eof or rs.bof then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 zbyh="<font color=red>["&mp&"]</font><font color=green>["&sf&"]</font><img src=../ico/"& touxiang &"-2.gif width=32 height=32>##想使用振臂一呼来提高自已的攻击力，可惜["&mp&"]下没有一个比##管理等级低的弟子，振臂一呼功能使用失败！"
 exit function
end if

mpzxrs=0
zbjwg=0
zbjgj=0
zbjfy=0 
'开始统计在线人员数目
do while not(rs.eof or rs.bof)
 name=rs("姓名")
 if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom"))),LCase(" "&name&" "))<>0 then
  mpzxrs=mpzxrs+1
  zbjwg=int(zbjwg+rs("武功")*wgxs)
  zbjgj=int(zbjgj+rs("攻击")*gjxs)
  zbjfy=int(zbjfy+rs("防御")*fyxs)
 end if
 rs.movenext
loop
rs.close
if mpzxrs=0 then
 set rs=nothing
 conn.close
 set conn=nothing
 zbyh="<font color=red>["&mp&"]</font><font color=green>["&sf&"]</font><img src=../ico/"& touxiang &"-2.gif width=32 height=32>##想使用振臂一呼来提高自已的攻击力，可惜["&mp&"]下没有一个比##管理等级低的弟子在线，振臂一呼功能使用失败！"
 exit function
end if
zb1=zb1+1
zb2=now()
zb3=zbjgj
zb4=zbjfy
zb5=zbjwg
bliu=zb1&"|"&zb2&"|"&zb3&"|"&zb4&"|"&zb5
conn.execute "update 用户 set 保留1='"& bliu &"',武功=武功+"&zbjwg&",攻击=攻击+"&zbjgj&",防御=防御+"&zbjfy&" where 姓名='"&sjjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
zbyh="<img src=../ico/"& touxiang &"-2.gif width=32 height=32>##向聊天室的弟子们大喝一声:“<font color=red><b>"&mp&"</b></font>的兄弟们，本门弟兄同生死，共进退，大家要同心协力，为我派扬威！”，"&mp&mpzxrs&"个在线弟子一呼百应，##武功立刻增加：<font color=red><b>"&zbjwg&"</b></font>，攻击增加：<font color=red><b>"&zbjgj&"</b></font>，防御增加：<font color=red><b>"&zbjfy&"</b></font>，有效时间为<font color=red><b>60分钟</b></font>。"
end function
%>