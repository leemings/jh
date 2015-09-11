<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=trim(request.querystring("id"))
if not isnumeric(id) then 
	Response.Write "<script language=JavaScript>{alert('提示：什么年代了？小子你还想来这招呀！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
id=trim(request.querystring("id"))
if InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where id=" & id & " and 介绍人='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('这个人也不是你介绍来的，你小子想作什么！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
toname=rs("姓名")
if rs("保留2")="扒皮"&Month(date()) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('call,你想作什么，你已经扒过皮了，还想作什么！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
yudian=rs("mvalue")
conn.Execute "update 用户 set 保留2='"&"扒皮"&Month(date())&"' where id="&id
conn.Execute "update 用户 set allvalue=allvalue+"&int(yudian*0.05)&" where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【扒点】：["&aqjh_name&"]从("&toname&")身上扒到泡点:"&int(yudian*0.05)&"点，努力吧，多拉人！</font>"		'聊天数据
says=replace(says,"'","\'")
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
Response.Write "<script Language=Javascript>alert('提示：["&aqjh_name&"]从("&toname&")身上扒到泡点:"&int(yudian*0.05)&"点，努力吧，多拉人！');location.href = 'javascript:history.go(-1)';</script>"
%>
