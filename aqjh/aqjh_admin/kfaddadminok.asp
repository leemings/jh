<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
kfnum=int(Request.Form("kfnum"))
kfnum=abs(kfnum)
hygrade=int(Request.form("hygrade"))

Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
Set rs=Server.CreateObject("ADODB.RecordSet")
if kfnum<1 or kfnum>1000 then
	Response.Write "<script language=javascript>alert('抱歉，一次加卡费必须在1~1000之间！请查看是否正确！');history.back();</script>"
	response.end
else
   sql="update 用户 set 会员金卡=会员金卡+"&kfnum&" where 会员等级="& hygrade 
   Set rs=conn.Execute(sql)
end if

set rs=nothing
conn.Close
set conn=nothing
says="<font size=2 color=red>【金卡发放】</font>恭喜站长发放"&hygrade&"级会员金卡"&kfnum&"元设置完成！"
act="消息"
towhoway=1
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
Response.Write "<script language=javascript>alert('恭喜你["&hygrade&"]级会员的会员金卡"&kfnum&"设置完成！');history.back();</script>"
%>