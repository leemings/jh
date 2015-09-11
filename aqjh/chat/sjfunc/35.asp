<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'宝物修练 xiulian
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(6)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间没有随机事件不可修练！');}</script>"
	Response.End
end if
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
'对系统禁止字符处理
if aqjh_grade<9 then
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
conn.open Application("aqjh_usermdb")
rs.open "select 宝物,宝物修练,等级,内力,操作时间,w5,门派 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("宝物")<>Application("aqjh_baowuname") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你并没有江湖宝物["& Application("aqjh_baowuname") &"]操作错误！');}</script>"
	Response.End
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<60 then
	s=60-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你还没有修练完成，请等["& s &"]秒再进行操作！');}</script>"
	Response.End
end if
if rs("内力")<2500  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('内力少于2500，无法修练！');}</script>"
	Response.End
end if
zstemp=add(rs("w5"),"养猪卡",1)
if rs("宝物修练")>=int(rs("等级")/Application("aqjh_baowuxl"))+1 then
	conn.execute "update 用户 set 宝物='无',法力=法力+2000,内力加=内力加+宝物修练,体力加=体力加+宝物修练,武功加=武功加+宝物修练,银两=银两+("& Application("aqjh_baowuyin") &"*宝物修练),w5='"&zstemp&"' where 姓名='" & aqjh_name &"'"
	xll=rs("宝物修练")
	conn.execute "update 用户 set 宝物修练=0,操作时间=now() where 姓名='" & aqjh_name &"'"
	xiulian="##<bgsound src=wav/xl.wav loop=1>祝贺您,您把你的的江湖宝物<font color=red>"& Application("aqjh_baowuname") &"</font>修练完成,所有上限加<font color=red>"& xll &"</font>点,得现金：+"& Application("aqjh_baowuyin")*xll &"两,法力上升了2000点,宝物修练完成，自动消失了。。。。。<font color=blue>哪料到从宝物中暴出<font color=red>1张养猪卡</font>，真神！~！。</font>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<60 then
	s=60-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你还没有修练完成，请等["& s &"]秒再进行操作！');}</script>"
	Response.End
end if
if rs("内力")<2500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('内力少于2500，无法修练！');}</script>"
	Response.End
end if
randomize timer
r=int(rnd*7)+1
select case r
	case 1 '金币
		n=int(rnd*1)+2
		conn.execute "update 用户 set 金币=金币+"&n&" where 姓名='" & aqjh_name &"'"
		lw="江洋大盗死于非命，遗留下<font color=brown>金币</font><font color=red>"&n&"</font>枚"
	case 2 '金卡
		n=int(rnd*1)+1
		conn.execute "update 用户 set 会员金卡=会员金卡+"&n&" where 姓名='" & aqjh_name &"'"
		lw="六色轮回，天惊地变，斗神赠送<font color=brown>会员金卡</font><font color=red>"&n&"</font>块"
	case 3 '种花卡
		n=int(rnd*1)+1
		temp=add(rs("w5"),"种花卡",n)
		conn.execute "update 用户 set w5='"&temp&"' where 姓名='" & aqjh_name &"'"
		lw="奇宝既奇宝，只见一绿光闪过，天上降下<font color=brown>种花卡</font><font color=red>"&n&"</font>张"
	case 4 '智力
		n=int(rnd*50)+1
		conn.execute "update 用户 set 智力=智力+"&n&" where 姓名='" & aqjh_name &"'"
		lw="白无常，黑无常，万物生灵降归附，再生诸葛亮传授<font color=brown>智力</font><font color=red>"&n&"</font>点"
	case 5 '知质
		n=int(rnd*10)+1
		conn.execute "update 用户 set 知质=知质+"&n&" where 姓名='" & aqjh_name &"'"
		lw="心动，情动，奇宝泛光，大涨<font color=brown>知质</font><font color=red>"&n&"</font>点"
	case 6 '养猪卡
		n=int(rnd*1)+1
		temp=add(rs("w5"),"养猪卡",n)
		conn.execute "update 用户 set w5='"&temp&"' where 姓名='" & aqjh_name &"'"
		lw="世上无奇不有，只见一雕闪过，天上降下<font color=brown>养猪卡</font><font color=red>"&n&"</font>张"
end select 
conn.execute "update 用户 set 宝物修练=宝物修练+1,内力=内力-2500,操作时间=now() where 姓名='" & aqjh_name &"'"
xiulian="##拥有江湖宝物"& Application("aqjh_baowuname") &"进行修练，这是你第:"& rs("宝物修练") & "次进行宝物修练....<font color=0088FF>爱情奇宝，世间罕有，每修炼一次暴一物品，</font><font color=blue>"&lw&"</font>!! "
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>