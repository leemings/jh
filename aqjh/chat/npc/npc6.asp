<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
st=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if st="" then Response.Redirect "../../error.asp?id=440"
if Application("aqjh_klt6")=0 then
	Response.Write "<script Language=Javascript>alert('提示：怪物已经不在了！');</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("aqjh_klt6"))))
'if tempjs<>r then
	'Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	'response.end
'end if
Application.Lock
Application("aqjh_klt6")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='"&st&"'"
rs.open "select 体力 FROM 用户 WHERE 姓名='"&st&"'",conn
if rs("体力")<0 then
	conn.execute "update 用户 set 状态='死',事件原因='怪物|"&fn1&"' where 姓名='"&st&"'"
conn.execute "insert into l(b,a,c,e,d) values ('" & st & "',now(),'吝啬','体力不足累死','人命')"
	kl="<img src='pic/kl.gif'>太不幸了，["&st&"]为了抵御怪物袭击，舍身取义，把生命永远的留在了彩虹之下，陪伴他的只有千百年仍在变幻着神奇光芒的传说………"
	Response.Write "<script Language=Javascript>alert('提示：你英勇可嘉，但是体力太弱，丢了小命，去复活吧！没能耐别乱闯！');</script>"
	call boot(st,"怪物，操作者：怪物，咆哮！")
else
	conn.execute "update 用户 set 魅力=魅力+"& tempjs *500&" where 姓名='" & st & "'"
	kl="<font color=#0000FF>"&st&"</font>一阵拳脚打跑寒冰卡比，并没杀死它，魅力"&tempjs*500&"点！"
	end if
	rs.close
set rs=nothing
conn.close
set conn=nothing

says="<font color=red><b>【打击怪物】</b></font>"&kl	
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
