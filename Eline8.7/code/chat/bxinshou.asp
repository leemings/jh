<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Application("sjjh_jinbi")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你真不够诚心，新手有人照了...');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
temp=int(abs(clng(Application("sjjh_jinbi"))))
if temp<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("sjjh_jinbi")=0
Application.UnLock
name=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.execute "update 用户 set 金币=金币+"& temp &",魅力=魅力+500 where 姓名='" & sjjh_name &"'"
jinbi="¤呵呵，能帮新朋友一把，有好处啊，大家也要多多照着一点嘛！因此：["&name&"]得到站长奖励的金币<img src='img/jinbi.gif'>"&temp&"个，魅力顿时上升500点！"
conn.close
set conn=nothing
says="<font color=red><b>【江湖消息】</b></font>"&jinbi		'聊天数据

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
