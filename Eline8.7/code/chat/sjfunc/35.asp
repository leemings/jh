<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'宝物修练♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(6)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间没有随机事件不可修练！');}</script>"
	Response.End
end if
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
says="<font color=red>【宝物修练】<font color=" & saycolor & ">"+xiulian()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'宝物修练 xiulian
function xiulian()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 宝物,宝物修练,等级,内力,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("宝物")<>Application("sjjh_baowuname") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你并没有江湖宝物["& Application("sjjh_baowuname") &"]操作错误！');}</script>"
	Response.End
end if
if rs("宝物修练")>=int(rs("等级")/Application("sjjh_baowuxl"))+1 then
	conn.execute "update 用户 set 宝物='无',内力加=内力加+宝物修练,体力加=体力加+宝物修练,武功加=武功加+宝物修练,银两=银两+("& Application("sjjh_baowuyin") &"*宝物修练) where 姓名='" & sjjh_name &"'"
	xll=rs("宝物修练")
	conn.execute "update 用户 set 宝物修练=0,操作时间=now() where 姓名='" & sjjh_name &"'"
	xiulian="##<bgsound src=wav/shuij.mid loop=1>祝贺您,您把你的江湖宝物<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname") &"</font>修练完成<img src=sjfunc/18.gif>,所有上限加<font color=red>"& xll &"</font>点,得现金：+"& Application("sjjh_baowuyin")*xll &"两,宝物修练完成，自动消失了。。。。。。"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<(int(rnd*120)+1) then
	s=(int(rnd*120)+1)-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你还没有修练完成，请等["& s &"]秒再进行操作！');}</script>"
	Response.End
end if
if rs("内力")<500  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('内力少于500，无法修练！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 宝物修练=宝物修练+1,内力=内力-500,操作时间=now() where 姓名='" & sjjh_name &"'"
xiulian="<bgsound src=wav/xlsj.mid loop=1>##拥有江湖宝物<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname") &"</font>进行<img src=sjfunc/17.gif>修练，这是你第:"& rs("宝物修练") & "次进行宝物修练。"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>