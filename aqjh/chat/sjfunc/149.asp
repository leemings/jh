<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'成为盟主
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=Session("aqjh_name")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
'房间信息
aqjh_roominfo=split(Application("aqjh_room"),";")
nowinroom=session("nowinroom")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
chatroomname=trim(chatinfo(0))
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
'房间信息
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
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【成为盟主】<font color=" & saycolor & ">"+mengzhu()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function mengzhu()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if chatroomname<>"决战爱情" then
 mengzhu="[##]申请盟主只有在[决战爱情]房间内进行！"
 exit function
end if
'判断房间在线
for i=0 to chatroomnum
 chatroomname=Application("aqjh_chatroomname"&i)
 if chatroomname="决战爱情" then
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
	chatinfo=split(aqjh_roominfo(nowinroom),"|")
	useronlinename=Application("aqjh_useronlinename"&nowinroom)
	if onlinenow<>1 then
		Response.Write "<script language=JavaScript>{alert('警告:战胜所有人后再来申请吧！');}</script>"
		Response.End
	end if
 end if
next
'判断时间
if Minute(time())<50 then
 		Response.Write "<script language=JavaScript>{alert('警告:时间还没结束就想领盟主令？');}</script>"
		Response.End
end if
if Application("aqjh_mengzhu")=aqjh_name then
 		Response.Write "<script language=JavaScript>{alert('警告:你已经是盟主了啊，还抢什么啊？');}</script>"
		Response.End
end if
conn.execute "update 用户 set 金币=金币+200 where 姓名='"&aqjh_name&"'"
Application("aqjh_mengzhu")=aqjh_name
mengzhu="恭喜[##]胜任江湖武林盟主的职位，官府送盟主令一个和200个金币！"
response.write "<script>alert('恭喜你抢得盟主令!');</script>"
conn.close
set conn=nothing
end function
%>