<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if InStr(Application("aqjh_diuqi"),"|")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：你来晚了……');}</script>"
	Response.End
end if
zt=split(Application("aqjh_diuqi"),"|")
'2银两
if ubound(zt)=2 then
	if zt(2)=aqjh_name then
		Response.Write "<script language=JavaScript>{alert('提示：自己丢弃的银两也要？！');}</script>"
		Response.End 
	end if
elseif ubound(zt)=4 then
	if zt(4)=aqjh_name then
		Response.Write "<script language=JavaScript>{alert('提示：自己丢弃的物品也要？！');}</script>"
		Response.End 
	end if
end if
Application.Lock
	Application("aqjh_diuqi")=""
Application.UnLock
if left(zt(0),1)<>"w" and left(zt(0),1)<>"v" then
	if not isnumeric(zt(0)) then 
		Response.Write "<script language=JavaScript>{alert('提示：数据有问题不可以操作！');}</script>"
		Response.End 
	end if
	yin=int(abs(clng(zt(0))))
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	conn.execute "update 用户 set 银两=银两+"&yin&" where  姓名='"&aqjh_name&"'"
	conn.close
	set conn=nothing
	diuqi="<font color=red><b>【江湖消息】</b></font>##真是幸运，得到了<font color=blue>["&zt(2)&"]</font>丢弃的银两[<font color=red><b>"&yin&"</b></font>]………"
else
if left(zt(0),1)="v" then
	if not isnumeric(zt(2)) then 
		Response.Write "<script language=JavaScript>{alert('提示：数据有问题不可以操作！');}</script>"
		Response.End 
	end if
	value=int(abs(clng(zt(2))))
	name=zt(1)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	conn.execute "update 用户 set allvalue=allvalue+"&value&" where  姓名='"&aqjh_name&"'"
	diuqi="<font color=red><b>【江湖消息】</b></font>##真是幸运，["&name&"]掉出聊天室，没有保存的存点：</b>"&value&"</b>点，让##抢到了…</font>]………"
	conn.close
	set conn=nothing
else
	if not isnumeric(zt(2)) then 
		Response.Write "<script language=JavaScript>{alert('提示：数据有问题不可以操作！');}</script>"
		Response.End 
	end if
	wusl=int(abs(clng(zt(2))))
	wpname=zt(1)
	lb=zt(0)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select "&lb&" from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
	zstemp=add(rs(lb),wpname,wusl)
	conn.execute "update 用户 set "&lb&"='"&zstemp&"' where  姓名='"&aqjh_name&"'"
	diuqi="<font color=red><b>【江湖消息】</b></font>##真是幸运，得到了<font color=blue>["&zt(4)&"]</font>丢弃的物品[<font color=red><b>"&wpname&"共："&wusl&"个</b></font>]………"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end if
zj="<a href=javascript:parent.sw('[" & aqjh_name & "]'); target=f2>"& aqjh_name & "</a>"
br="<a href=javascript:parent.sw('[" & name & "]'); target=f2>" & name & "</a>"
diuqi=Replace(diuqi,"##",zj,1,3,1)
diuqi=Replace(diuqi,"%%",br,1,3,1)
says=diuqi
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
