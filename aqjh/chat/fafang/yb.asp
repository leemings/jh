<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/func.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Application("aqjh_yb")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子来晚了，元宝没有了...');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
temp=int(abs(clng(Application("aqjh_yb"))))
if temp<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("aqjh_yb")=0
Application.UnLock
name=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 银两=银两+"& temp &" where 姓名='" & aqjh_name &"'"
yb="今天真是走运，出门都有钱砸到我的脑袋上，呵。。。因此：["&name&"]得到:<img src='img/251.gif'>"&temp&"两!"
conn.close
set conn=nothing
says="<font color=red><b>【江湖消息】</b></font>"&yb		'聊天数据
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
