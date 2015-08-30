<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'宝物修练♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
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
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【紫金葫芦】<font color=" & saycolor & ">"+xiulian()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'宝物修练 xiulian
function xiulian()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select * from 夺宝",conn
slid=rs("冠军")
xlcs=rs("修练次数")
xb=rs("宣布")
lq=rs("领取")
rs.close
if xb=false then	'如果未宣布则不能修练
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('夺宝大赛结果还未出来，紫金葫芦不是你的！');}</script>"
	Response.End
end if
if lq=true then	'如果已领取宝物则不能修练
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('紫金葫芦都已经修练完了，你还修练个什么劲！');}</script>"
	Response.End
end if

rs.open "select id,等级,体力,内力,银两,操作时间 FROM 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("id")<>slid then	'如果非冠军不能修练
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是夺宝大赛的冠军，紫金葫芦不是你的！');}</script>"
	Response.End
end if

cs=int(rs("等级")/40)+1	'修练所需要的次数
if xlcs>=cs then	'如果修练次数够了，则不能修练
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已经修练完紫金葫芦了！');}</script>"
	Response.End
end if
sxyl=(xlcs+1)*10000000
if rs("银两")<sxyl then	'修练时所需银两
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是第"&xlcs+1&"次修练宝物，需要"&sxyl&"银两，你的钱不够！');}</script>"
	Response.End
end if
if cs<=(xlcs+1) then	'是否为最后一次修练
	conn.execute "update 夺宝 set 修练次数="& cs &",领取=true"
	conn.execute "update 用户 set 银两=银两-"& sxyl &",金币=金币+500,allvalue=allvalue+3000,内力加=内力加+1000,体力加=体力加+1000,武功加=武功加+1000,操作时间=now() where 姓名='" & sjjh_name &"'"
	xiulian="<bgsound src=wav/xlw.mid loop=1>##祝贺您，夺宝大赛奖品：<font color=red><b>紫金葫芦</b></font>修练完成<img src=sjfunc/18.gif>，你的<font color=red><b>体力、内力、武功上限各加1000点，得到500个金币，经验增加3000点</b></font>，不知下届夺宝大赛冠军得主是。。。"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<180 then
	s=180-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你还没有修练完成，请等["& s &"]秒再进行操作！');}</script>"
	Response.End
end if
if rs("内力")<5000 or rs("体力")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('体力或内力少于5000，无法修练！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 内力=内力-5000,体力=体力-5000,银两=银两-"&sxyl&",操作时间=now() where 姓名='" & sjjh_name &"'"
conn.execute "update 夺宝 set 修练次数=修练次数+1"
xiulian="<bgsound src=wav/xlhl.mid loop=1>##夺得江湖夺宝大赛冠军，得到江湖至宝<font color=red><b>紫金葫芦</b></font>，呵呵，有宝贝就是好！这是你第<font color=#000000><b>"& xlcs+1 & "</b></font>次进行宝物<img src=sjfunc/xlhl.gif>修练。"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>