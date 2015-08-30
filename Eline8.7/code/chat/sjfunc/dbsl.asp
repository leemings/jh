<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="chatconfig.asp"-->
<%'夺宝胜利宣告♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
sjjh_name=session("sjjh_name")
sjjh_grade=session("sjjh_grade")
inroom=session("nowinroom")
call roompd("胜利公告")
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=0
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=slgg(mid(says,i+1),inroom)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'公告
function slgg(fn1,inroom)
	sjjh_roominfo=split(Application("sjjh_room"),";")
	chatinfo=split(sjjh_roominfo(inroom),"|")
	fjname=chatinfo(0)
	erase sjjh_roominfo
	erase chatinfo
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select * from 夺宝",conn,2,2
	xb=rs("宣布")
	slid=rs("冠军")
	rs.close
	if xb then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：胜利者或系统已经向大家宣布过夺宝胜利公告了！');}</script>"
		response.end
	end if
	zxrs=application("sjjh_onlinelist"&inroom)
	 '开始统计在线人员数目
	rjsq=ubound(zxrs)
	jsq=0
	sname=""
	for i=1 to rjsq
		sfgf=split(zxrs(i),"|")
		if sfgf(0)<>session("sjjh_name") and sfgf(2)<>"官府" then
			sname=sfgf(0)
			jsq=jsq+1
		end if
		if jsq>1 then
			erase zxrs
			erase sfgf
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('提示：比赛还未结束，还有除你以外的非官府人员在夺宝大赛房间！');}</script>"
			response.end
		end if
		erase sfgf
	next
	erase zxrs
	if sname="" then sname=sjjh_name
	rs.open "select id from 用户 where 姓名='"&sname&"'",conn
	if not (rs.eof or rs.bof) then
		slid=rs("id")
		rs.close
		conn.execute "update 夺宝 set 冠军="& slid &",时间=now(),领取=false,修练次数=0,宣布=true"
	else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：比赛最后胜利者："&sname&"的账号在江湖中不存在，请于站长联系！');}</script>"
		response.end
	end if
	set rs=nothing
	conn.close
	set conn=nothing
	gg="<font color=blue>"& sname &"</font>经过艰苦奋战，终于夺得了<font color=red><b>夺宝大赛冠军</b></font>，站长"& application("sjjh_user") &"将武林至宝<font color=red><b>紫金葫芦</b></font>奖励给<font color=blue>"& sname &"</font><img src=sjfunc/guanjun.gif>，此宝需要修练可提升自已的各种上限及得到金币。"
	slgg="<bgsound src=wav/tsbj.mid loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>夺宝大赛胜利公告</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
	call dbxx(slgg)
end function
%>