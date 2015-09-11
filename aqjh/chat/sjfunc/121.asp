<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'分喜糖
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))

'对暂离开、点哑穴判断
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
says="<font color=green>【发喜糖】<font color=" & saycolor & ">"+fen(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'分喜糖
function fen(fn1,to1)
fn1=int(abs(fn1))
if fn1<500 or fn1>1000 then
	Response.Write "<script language=JavaScript>{alert('提示：500-1000两');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两,配偶,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("银两")<50000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的钱分给大家吗?');}</script>"
	Response.End
end if
if rs("配偶")="无" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你还没结婚,分什么喜糖啊！');}</script>"
	Response.End
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等"& s &"秒再操作,可别累着！');</script>"
	response.end
end if

Response.Buffer=true
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
for i=0 to x
conn.Execute "update 用户 set 银两=银两+" & fn1 & " where 姓名='" & online(i) & "'"
next
conn.Execute "update 用户 set 银两=银两-" & fn1*1000 & ",道德=道德+1000,操作时间=now() where 姓名='" & aqjh_name & "'"
fen="##<bgsound src=wav/FAQIAN.wav loop=1>与二人["&rs("配偶")&"]在这个大喜的日子,拿出"&fn1*1000&"两给大家发喜糖!每人分得银两"&fn1&",道德增加1000,大家恭喜他(她)们了!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
