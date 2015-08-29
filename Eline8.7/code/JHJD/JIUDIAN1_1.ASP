<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
id=request("id")
%><!--#include file="dadata.asp"-->
<%
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能订购酒宴"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "SELECT * FROM 宴会 where ID=" & id,conn
wu=rs("宴会名")
yin=rs("售价")
lx=rs("类型")
nl=rs("内力")
tl=rs("体力")
jb=rs("金币")
sl=rs("数量")%>
<%
rs.close
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if yin>=rs("银两") then
	rs.close
	set rs=nothing
	connt.close
	set connt=nothing
	conn.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('不能定酒宴，原因：你的银两不够了');history.back();</script>"
	response.end
end if
	conn.execute "update 用户 set 银两=银两-" & yin & " where 姓名='" & sjjh_name & "'"
	rs.close
	rs.open "select * from 宴会列表 where 宴会名='" & wu & "' and 拥有者='" & sjjh_name & "'",conn
if rs.eof or rs.bof then
	connt.execute "insert into 宴会列表(宴会名,拥有者,类型,内力,体力,金币,数量,售价,时间) values ('"&wu&"','"&sjjh_name&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now())"
	rs.close
	set rs=nothing
	connt.close
	set connt=nothing
says="<font color=red>【好消息】</font><b><font color=green>"&sjjh_name&"在江湖大酒店的"&wu&"厅举行<font color=red>※"&lx&"※</font>宴会，大家都快去呀，晚了就没的吃了。。。</font></b>"			'聊天数据
says=replace(says,"'","\'")
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
	Response.Redirect "jd.asp"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('不能定酒宴，原因：你已定了同样规格的酒席！');history.back();</script>"
	response.end
end if
%>
