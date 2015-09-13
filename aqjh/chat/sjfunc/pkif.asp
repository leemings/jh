<%
'对贵宾的判断
toname=Trim(Request.Form("towho"))
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：你是本江湖贵宾,不得参与任何恩怨！');}</script>"
	Response.End
end if
if Instr(Application("aqjh_guibin"),"|" & toname & "|")<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：对方是本江湖贵宾！');}</script>"
	Response.End
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
'对发招的判断
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以发招！');}</script>"
	Response.End
end if
'对pk时间的判断
f=Minute(time())
pktime=cint(chatinfo(10))
if f>pktime then
        Response.Write "<script language=JavaScript>{alert('提示：江湖PK打架时间为每小时前["&pktime&"]分！');}</script>"
	response.end
end if
'对决战房的判断
if chatinfo(0)="决战快乐" and Hour(time())=18 then
        Response.Write "<script language=JavaScript>{alert('提示：决战房PK时间未到！');}</script>"
	response.end
end if
%>