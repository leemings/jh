<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("boy")="" and rs("boysex")="" then
   	Response.Write "<script Language=Javascript>alert('提示：你并没有小孩啊？是做梦吧！');history.back();</script>"
	response.end
end if
if rs("道德")<10000 then
	Response.Write "<script Language=Javascript>alert('提示：你的道德那已经很低了，你就积点德吧！');history.back();</script>"
	response.end
else
 conn.execute "update 用户 set 道德=道德-10000,boy='',boysex='' where 姓名='" & aqjh_name &"'"
 kl=aqjh_name&"把自己的婴儿丢在了路边，依依不舍地离去，损失道德<font color=red>10000</font>点."
 Response.Write "<script Language=Javascript>alert('唉！真是怨孽啊！');history.back();</script>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【江湖怨孽】</b></font>"&kl
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