<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
response.buffer=true
'=========================================================
' File: events.asp
' For DVbbs6.0
' Date: 2003-4-20
' Script Written by FSSunwin
'=========================================================
' Copyright (C) 2001,2002 artbbs.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net,http://www.artbbs.net,http://artbbs.3322.org
' Email: sunwin2000@yeah.net
'=========================================================

dim mlev,strSql,strCMessage,MemberErased
dim sEName
dim abgcolor
dim usersign
dim USERNAME
usersign=false
redim Board_Setting(4)
Board_Setting(4)=1 '发帖模式为高级

stats="论坛日记"
call nav()
call head_var(2,0,"","")

dim intPrivateEvent
dim intThisMonth,intThisYear,ddate
dim sSQL, sStart_Date, sEnd_Date, sEvent_title, sEvent_Details, sMode, berror,eventid
dim eMode
dim datToday
Const cSUN = 1, cMON = 2, cTUE = 3, cWED = 4, cTHU = 5, cFRI = 6, cSAT = 7
datToday = date()
	'判断日记的属性 
	if Request.Form("isPrivateEvent") = "1" then
		intPrivateEvent = 1
	else
		intPrivateEvent = 0
	end if
	
	intThisMonth = Request.QueryString("month")
	if intThisMonth = "" then
		intThisMonth = month(datToday)
	else
		intThisMonth = CInt(intThisMonth)
	end if
	If intThisMonth < 1 OR intThisMonth > 12 Then
		intThisMonth = Month(datToday)
	End If
	
	intThisYear = Request.QueryString("year")
	If intThisYear = "" Then
		intThisYear = Year(datToday)
	else
		intThisYear = CInt(intThisYear)
	End If
	eMode = Trim(Request.QueryString("eMode"))
	'if eMode="" or isnull(eMode) then eMode="年历"
	sMode = Trim(Request.QueryString("mode"))

	dDate = Request.QueryString("Date")
	If not IsEmpty(ddate) and IsDate(ddate) and smode = "" Then
		sMode = "display"
	end if

	if Request.Form("EVENT") <> "" then
		Update_Event(Request.Form("EVENT"))
	end if	

	select case sMode
		case "delid"
			if founduser then
				if master then
				sSQL = "DELETE * FROM EVENTS WHERE Event_ID=" & cint(Request.QueryString("Event_ID"))
				else
				sSQL = "DELETE * FROM EVENTS WHERE adduser='"&checkstr(trim(Request.QueryString("username")))&"' and Event_ID=" & cint(Request.QueryString("Event_ID"))
				end if
			conn.Execute ssql

			strCMessage = "所选择的帖子已被删除！"
			else
			Errmsg=Errmsg + "<li>删除操作失败,您未登陆或不是日记作者！"
			founderr=true
			end if
		case "edit"
			if founduser then
			sMode = "edit"
			strCMessage = "修改成功！"
				else
			Errmsg=Errmsg + "<li>修改操作失败,您未登陆或不是日记作者！"
			founderr=true
			end if
		case "add"
			if founduser then
			smode = "add"
				else
			Errmsg=Errmsg + "<li>请您先登陆才能添写日记！"
			founderr=true
			end if
	end select

	if founderr then 
	call dvbbs_error()
	call footer()
	response.end
	end if

	dim strMonthName, datFirstDay, intFirstWeekday, intLastDay, intPrevMonth, intNextMonth, intPrevYear, intNextYear
	dim LastMonthDate, NextMonthDate, intPrintDay, intLastMonth, dToday, dFirstDay, dLastDay, endrows, intLoopDay
	dim bevents,  sTitle

%>
<link rel=stylesheet type="text/css" href="inc/events.css">
<script Language="JavaScript">
function showlist(url, width, height){
	var Win = window.open(url,"showlist",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
</script>
<TABLE cellPadding=0 cellSpacing=1 align=center class=tableborder1>
<tr>
<th height="24" >栏目导航
</th>
<th ><a href="<%=ScriptName%>?emode=年历&year=<%=intThisYear%>" >显示年历</a>
</th>
<th ><a href="<%=ScriptName%>?year=<%=intThisYear%>&month=<%=intThisMonth%>" >显示月历</a>
</th>
<th ><a href="<%=ScriptName%>?emode=月历&year=<%=intThisYear%>&month=<%=intThisMonth%>" >显示日记</a>
</th>
<th ><a href="<%=ScriptName%>?showmode=birthday&year=<%=intThisYear%>&month=<%=intThisMonth%>" >生日列表</a>
</th>
</tr>
	<tr>
		<td  align="center" valign="top" nowrap width="140" class=tablebody1>
			<table border="0" cellpadding="1" cellspacing="2" width="140" height="100%">
				<%emitmonths()%>
				<tr>
					<td align=center>
						<form ACTION="<%=ScriptName%>" METHOD=GET id=frmSelectMonth name=frmSelectMonth>
							<select NAME="year">
								<%for i = -3 to 3%>
								<OPTION VALUE="<%= intThisYear + i%>" <%if (intThisYear + i) = intThisYear then Response.Write(" Selected")%>><%= intThisYear + i%>
								<%next%>
							</select>&nbsp;年
							<select NAME="month">
								<%for i = 1 to 12%>
								<OPTION VALUE="<%=i%>" <%if i=intThisMonth then Response.Write("Selected")%>><%=i%>
								<%next%>
							</select>&nbsp;月<br><br>
							<input TYPE="submit" NAME="Go" value=".显 示." BORDER="0" WIDTH="35" HEIGHT="20">
						</form>
					</td>
				</tr>
				<%if founduser then%>
				<tr>
					<td align=center>欢迎您：<FONT color=<%=Forum_body(8)%>><%= membername %></FONT><BR>
						『<a href="z_events.asp?mode=add&date=<%if dDate= "" then Response.Write(datToday) else Response.Write(dDate)%>">新增日记</a>』
					</td>
				</tr>
				<%end if
				  if smode <> "" then%>
				<tr>
					<td align=center>
						<div align="center"><a href=z_events.asp>回到月历首页</a></div>
					</td>
				</tr>
				<%end if%>
				<tr><td>&nbsp;</td></tr>
				<tr><th align="center" class=tabletitle2>
					当前日记：
				</th></tr>
				<tr>
				<td>
				<% call emitupcomingevents %>
					<br><br>
					</td>
				</tr>
				<tr>
					<th align="center" >
					过去日记：</th>
				</tr>
				<tr>
					<td>
					<% call emitpastEvents %>
					<br>
					<br>
					</td>
				</tr>
			</table>
		</td>
		<td class=tablebody1 align="center" valign="middle" colspan="4" >
		<%if smode = "" and ( eMode="" or eMode="日历") then
			emitmonths2
			elseif eMode="年历"	then
		%>
        <div align="justify">
		<table cellpadding=0 cellspacing=0 align=center border="0" width="100%">
		<tr>
<% 
dim mon ,tr
tr=0
for mon=1 to 12 
tr=tr+1
%>
<td height=140 Valign=top>
<%call mondate(mon,1)%>
</td>
<% if tr>=3 then 
response.write "</tr>" 
tr=0
end if
next
%>
</tr></table></div>
		<%else%>
			<TABLE cellPadding=0 cellSpacing=1 align=center class=tableborder2 style="width:100%">
				<tr>
				<td class=tablebody2 align="center" valign="middle" nowrap>
				<% call emitEvents()%>
				</td>
				</tr>
				<tr>
				<td>
				<< <a href=z_events.asp>返回月历首页</a>
				</td>
				</tr>
				</table>
				<%end if%>	
		</td>
	</tr>
<tr>
<th height="24" >栏目导航
</th>
<th ><a href="<%=ScriptName%>?emode=年历&year=<%=intThisYear%>" >显示年历</a>
</th>
<th ><a href="<%=ScriptName%>?year=<%=intThisYear%>&month=<%=intThisMonth%>" >显示月历</a>
</th>
<th ><a href="<%=ScriptName%>?emode=月历&year=<%=intThisYear%>&month=<%=intThisMonth%>" >显示日记</a>
</th>
<th ><a href="<%=ScriptName%>?showmode=birthday&year=<%=intThisYear%>&month=<%=intThisMonth%>" >生日列表</a>
</th>
</tr>
</table>
<%
'------------------------------------------------------------
' 当月最后一天日期函数
'------------------------------------------------------------
Function GetLastDay(intMonthNum, intYearNum)
	Dim dNextStart
	If CInt(intMonthNum) = 12 Then
		dNextStart = CDate( "1/1/" & intYearNum)
	Else
		dNextStart = CDate(intMonthNum + 1 & "/1/" & intYearNum)
	End If
	GetLastDay = Day(dNextStart - 1)
End Function
	
'-------------------------------------------------------------------------
' 左边小日历显示
'-------------------------------------------------------------------------
Sub Write_TD(sValue, sClass)
	Response.Write "		<td ALIGN='RIGHT' WIDTH=20 HEIGHT=15 VALIGN='BOTTOM' CLASS='" & sClass & "'> " & sValue & "</td>" & vbCrLf
End Sub

'-------------------------------------------------------------------------
' 显示该天日记数据
'-------------------------------------------------------------------------

Sub Write_TD2(sValue, sClass, dDate)
	dim rsEvents, sETitle
	set rsevents = server.CreateObject("adodb.recordset")
	Response.Write "		<td HEIGHT='80' WIDTH='14%' ALIGN='left' VALIGN='top' CLASS='" & sClass & "'> " & sValue
	sSQL = 	"SELECT top 5 event_id, start_date, end_date, event_title, event_details, adduser,PRIVATE FROM EVENTS " & _
			"WHERE Start_Date <= '" & DateTostr(dDate) & "' AND End_date >= '" & DateTostr(dDate) & "' ORDER BY Event_ID DESC"
	rsEvents.Open sSQL,conn,1,1
	do while not(rsevents.eof)
		sETitle = rsEvents("event_title")
		sEName = lcase(rsEvents("adduser"))
		if len(sETitle) > 14 then
			sETitle = mid(sETitle,1,15)
		end if
		if (rsEvents("PRIVATE") <> 1) or (rsEvents("PRIVATE") = 1 and sEName = lcase(membername)) and founduser then
			Response.Write "<br><FONT color=#b70000>·</FONT><a href=z_events.asp?date="& Server.URLEncode(dToday) & ">" & sETitle & "</a>"
		end if
		rsEvents.MoveNext
	loop
	Response.Write "</td>" & vbCrLf
	rsEvents.Close
	set rsevents = nothing
End Sub

'-------------------------------------------------------------------------
' 该天没有日记数据
'-------------------------------------------------------------------------

Sub Write_TD3(sValue, sClass)
	Response.Write "		<td HEIGHT='80' WIDTH='14%' ALIGN='left' VALIGN='top' CLASS='" & sClass & "'> " & sValue & "</td>" & vbCrLf
End Sub

'-------------------------------------------------------------------------
' 该天会员生日
'-------------------------------------------------------------------------
Sub showbday(sValue, sClass, dDate)
	dim srs,sSql
	dim BDate,birth,fonts
	BDate=split(dDate,"-")
	birth=BDate(1)&"-"&BDate(2)
	Response.Write "		<td HEIGHT='80' WIDTH='14%' ALIGN='left' VALIGN='top' CLASS='" & sClass & "'> " & sValue
	set srs = server.CreateObject("adodb.recordset")
	sSQL =  "select userid,username,birthday from [user] "
	sSQL = sSQL + "where birthday like '%"&birth&"%' "
	srs.Open sSQL,conn,1,1
	do while not(srs.eof)
	if month(srs(2))=intThisMonth and day(srs(2))=day(dDate) then
		if srs(1)=membername then fonts="<font color="&Forum_body(8)&" >"
	Response.Write "<br><FONT color="&Forum_body(8)&">·</FONT><a onMouseOver=""ShowMenu('<a style=font-size:9pt;line-height:14pt; href=dispuser.asp?id="&srs(0)&">查看资料</a><br><a style=font-size:9pt;line-height:14pt; href=messanger.asp?action=new&touser="&srs(1)&">给他留言</a>',50)"" > "&fonts+srs(1)&"</font></a>"
	end if	
	srs.MoveNext
	loop
	Response.Write "</td>" & vbCrLf
	srs.close
	set srs=nothing
End Sub

Sub Show_Form()
	Dim sButtonMsg
	If sMode = "edit" Then
		sButtonMsg = "更新日记"
	Else
		sButtonMsg = "新增日记"
	End If
	
	If bError = True Then
		sTitle = "送出资料有误"
	Else
		sTitle = "新增或修改日记"
	End If
%>
<TABLE cellPadding=0 cellSpacing=1 align=center class=tableborder1 STYLE="width:100%">
<form  ACTION="<%=ScriptName%>" METHOD="POST" name="frmAnnounce">
<tr>
<td class=tablebody2 VALIGN="TOP" ALIGN="right" WIDTH="15%">开始日期：</td>
<td class=tablebody2 VALIGN="TOP" ALIGN="LEFT" WIDTH="85%" ><input TYPE="TEXT" SIZE="12" MAXLENGTH="12" NAME="START_DATE" VALUE="<%if HTMLEncode(sStart_Date) <> "" then Response.write HTMLEncode(sStart_Date) else Response.Write(dDate) end if%>">
<input type="Checkbox" name="isPrivateEvent" value="1" <% if intPrivateEvent = 1 then response.write "checked" %>>私人日记</td>
</tr>
<tr>
<td class=tablebody2 VALIGN="TOP" ALIGN="RIGHT" >结束日期：</td>
<td class=tablebody2 VALIGN="TOP" ALIGN="LEFT" >
<input TYPE="TEXT" SIZE="12" MAXLENGTH="12" NAME="END_DATE" VALUE="<%if HTMLEncode(sEnd_Date) <> "" then Response.write HTMLEncode(sEnd_Date) else Response.Write(dDate) end if%>">
<select size=1 name=Weather>
                <option value=晴天 selected>晴天</option>
                <option value=阴天 >阴天</option>
                <option value=多云 >多云</option>
                <option value=有雾 >有雾</option>
                <option value=小雨 >小雨</option>
                <option value=中雨 >中雨</option>
                <option value=雷雨 >雷雨</option>
                <option value=酷热 >酷热</option>
                <option value=寒冷 >寒冷</option>
                <option value=清爽 >清爽</option>
                <option value=月圆 >月圆</option>
                <option value=月缺 >月缺</option>
</select>
</td></tr>
<tr>
<td class=tablebody2 VALIGN="TOP" ALIGN="RIGHT">日记标题：</td>
<td class=tablebody2 VALIGN="TOP" ALIGN="LEFT">
<%	if sEvent_Title<>"" then
		sEvent_Title=replace(sEvent_Title,"<br>",chr(13))
		sEvent_Title=replace( sEvent_Title,"&nbsp;"," ")
		end if
%>
<input TYPE="TEXT" SIZE="30" MAXLENGTH="100" NAME="EVENT_TITLE" VALUE="<%=sEvent_Title%>"></td>
</tr>
<tr>
<td class=tablebody2  ALIGN="RIGHT">高级功能：
</td>
<td class=tablebody2 VALIGN="TOP" ALIGN="LEFT">
<script src="inc/ubbcode.js"></script>
<%if Cint(Board_Setting(4)) then%>
<!--#include file="INC/getubb.asp"-->
<%end if%>
</td>
</tr>
<tr>
<td class=tablebody2 VALIGN="TOP" ALIGN="RIGHT">日记内容：</td>
<td class=tablebody2 VALIGN="TOP" ALIGN="LEFT">
<textarea COLS="70" ROWS="15" NAME="Content" wrap=VIRTUAL>
<%
	if sEvent_Details<>"" then
	sEvent_Details=replace(sEvent_Details,"<br>",chr(13))
	sEvent_Details=replace(sEvent_Details,"&nbsp;"," ")
	response.write sEvent_Details
	end if
%>
</textarea>
</td>
</tr>
<tr>
<td class=tablebody2>&nbsp;<% =Request.QueryString("EVENT_ID") %></td>
<td class=tablebody2>
<input TYPE="hidden" NAME="event" VALUE="<%if sMode = "add" then Response.Write "ADD" else Response.write sMode %>">
<input TYPE="hidden" NAME="event_id" VALUE="<% =Request.QueryString("EVENT_ID") %>">
<input TYPE="SUBMIT" VALUE="<% =sButtonMsg %>" id="SUBMIT1" name="SUBMIT1">
</td></tr></form></table>

<%
End Sub

Sub Update_Event(sUpdateMode)
	dim enddate, memberid
	bError = False
	if not founduser then
  	ErrMsg=ErrMsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
	else
	memberid = userid
	end if
	if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
	FoundErr=True
	end if

	if FoundErr=True then
		call dvbbs_error()
		exit sub
	end if

	sStart_Date = Request.Form("START_DATE")
	sEnd_Date = Request.Form("END_DATE")
	sEvent_Title = Checkstr(trim(Request.Form("EVENT_TITLE")))
	sEvent_Details = Checkstr(trim(Request.Form("Content")))
	
	If NOT IsDate(sStart_Date) Then
		bError = True
	End If
	
	If Trim(sEnd_Date) <> "" AND NOT IsDate(sEnd_Date) Then
		bError = True
	End If
	
	If Trim(sEvent_Title) = "" Then
		bError = True
	End If
	
	If Trim(sEvent_Details) = "" Then
		bError = True
	End If
	
	If IsDate(sStart_Date) AND IsDate(sEnd_Date) Then
		If CDate(sStart_Date) > CDate(sEnd_Date) Then
			bError = True
		End If
	End If

	If bError = False Then
		If Trim(sEnd_Date) <> "" Then
		enddate = CDate(Request.Form("END_DATE"))
		Else
		enddate = CDate(Request.Form("START_DATE"))
		End If
		If sUpdateMode = "edit" Then
			sSQL = "UPDATE EVENTS SET Start_Date = '" & DateToStr(sStart_Date) & "', End_Date = '" & DateToStr(enddate) & "', Event_Title = '" & sEvent_Title & "', Event_Details = '" & sEvent_Details & "',PRIVATE = " & intPrivateEvent & ", date_added = '" & DateToStr(now()) & "' WHERE Event_ID=" & clng(Request.Form("Event_ID"))
			conn.Execute sSQL
		Else
			sSQL = "INSERT INTO EVENTS (Start_Date, End_Date, Event_Title, Event_Details, Date_Added, Added_by, PRIVATE, adduser, weather) Values ('" & DateToStr(sStart_Date) & "','" & DateToStr(enddate) & "',' " & sEvent_Title & "','" & sEvent_Details & "','" & DateToStr(now()) & "', '" & memberid & "','"  & intPrivateEvent & "','"&checkstr(membername)&"','" & checkstr(request.Form("weather")) & "' )"
			conn.Execute sSQL
			response.write ssql
		End If
		If sUpdateMode = "ADD" Then
			sTitle = "日记已加入<br>感谢你的使用！"
		Else
			sTitle = "日记已更新<br>感谢你的使用！"
		End If
	Else
		If sUpdateMode = "ADD" Then
			sTitle = "<font color="&Forum_body(8)&">发生问题</font><br>"
			sTitle = sTitle & "日记并未加入<br>请重试一遍！"
		Else
			sTitle = "<font color="&Forum_body(8)&">发生问题</font><br>"
			sTitle = sTitle & "日记并未修改<br>请重试一遍！"
		End If
	End If	
End Sub	

Function emitmonths()
		strMonthName = MonthName(intThisMonth)
		datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
		intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
		intLastDay = GetLastDay(intThisMonth, intThisYear)
		
		' Get the previous month and year
		intPrevMonth = intThisMonth - 1
		If intPrevMonth = 0 Then
			intPrevMonth = 12
			intPrevYear = intThisYear - 1
		Else
			intPrevYear = intThisYear	
		End If
		
		' Get the next month and year
		intNextMonth = intThisMonth + 1
		If intNextMonth > 12 Then
			intNextMonth = 1
			intNextYear = intThisYear + 1
		Else
			intNextYear = intThisYear
		End If

		' Get the last day of previous month. Using this, find the sunday of
		' last week of last month
		if intThisMonth = "" then
			intLastMonth = DatePart( "m", DateAdd( "m", -1, datToday))
		else
		   	if intThisMonth = 1 then
		   		intLastMonth = 12
			else
				intLastMonth = intThisMonth - 1
			end if
		end if	
		if intThisYear = "" then
			intPrevYear = DatePart( "yyyy", DateAdd( "m", -1, datToday))
		else
			if intThisMonth = 1 then
		  		intPrevYear = intThisYear - 1
			else
		  		intPrevYear = intThisYear
			end if
		end if

		' Get the last day of previous month. Using this, find the sunday of
		' last week of last month
		LastMonthDate = GetLastDay(intLastMonth, intPrevYear) - intFirstWeekDay + 2
		NextMonthDate = 1

		' Initialize the print day to 1
		intPrintDay = 1

		' These dates are used in the SQL
		dFirstDay = intThisMonth & "/1/" & intThisYear
		dLastDay 	= intThisMonth & "/" & intLastDay & "/" & intThisYear

		sSQL = 	"SELECT event_id, start_date, end_date, event_title, event_details, adduser, Private FROM EVENTS  WHERE " &_
						"(Start_Date >='" & DateToStr(dFirstDay) & "' AND Start_Date <= '" & DateToStr(dLastDay) & "') " & _
						"OR " & _
						"(End_Date >='" & DateToStr(dFirstDay) & "' AND End_Date <= '" & DateToStr(dLastDay) & "') " & _
						"OR " & _
						"(Start_Date < '" & DateToStr(dFirstDay) & "' AND End_Date > '" & DateToStr(dLastDay) & "' )"  & _
						"ORDER BY Start_Date"
		set rs = server.CreateObject("adodb.recordset")
		rs.Open sSQL,conn,1,1
%>
	<tr>
	<td valign="top">
	<table ALIGN="CENTER" BORDER="1" CELLSPACING="1" CELLPADDING="4" BGCOLOR="White" BORDERCOLOR="Gray">
	<tr><td>
		<table WIDTH="140" BORDER="0" CELLPADDING="1" CELLSPACING="0" BGCOLOR="#FFFFFF">
			<tr HEIGHT="18" BGCOLOR="Silver">
				<td WIDTH="20" HEIGHT="18" ALIGN="LEFT" VALIGN="MIDDLE"><a href="<% =ScriptName%>?month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>"><img src="<%=Forum_info(7)%>prev.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="上一个月"></a></td>
				<td WIDTH="120" COLSPAN="5" ALIGN="CENTER" VALIGN="MIDDLE" CLASS="SOME"><% = strMonthName & " " & intThisYear %></td>
				<td WIDTH="20" HEIGHT="18" ALIGN="RIGHT" VALIGN="MIDDLE"><a href="<%=ScriptName%>?month=<% =IntNextMonth %>&amp;year=<% =IntNextYear %>"><img src="<%=Forum_info(7)%>next.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="下一个月"></a></td>
			</tr>
		  <tr>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">S</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">M</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">T</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">W</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">T</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">F</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">S</td>
		  </tr>
		  <tr><td HEIGHT="1" ALIGN="MIDDLE" COLSPAN="7"><img src="<%=Forum_info(7)%>line.gif" HEIGHT="1" WIDTH="140" BORDER="0"></td></tr>
		  <%
				EndRows = False
				Response.Write vbCrLf
				
			 	Do While EndRows = False
					Response.Write "	<tr>" & vbCrLf
					For intLoopDay = cSUN To cSAT
						If intFirstWeekDay > cSUN Then
							Write_TD LastMonthDate, "NON"
							LastMonthDate = LastMonthDate + 1
							intFirstWeekDay = intFirstWeekDay - 1
						Else
							If intPrintDay > intLastDay Then
								Write_TD NextMonthDate, "NON"
								NextMonthDate = NextMonthDate + 1
								EndRows = True
							Else
								If intPrintDay = intLastDay Then
									EndRows = True
								End If
								dToday = CDate(intThisMonth & "/" & intPrintDay & "/" & intThisYear)
								If NOT Rs.eof Then
								bEvents = False
								  Do While NOT Rs.eof AND bEvents = False
								  sEName = lcase(RS("adduser"))
								    If dToday >= strToDate(Rs("Start_Date")) AND dToday <= strToDAte(Rs("End_Date")) AND ((Rs("PRIVATE") <> 1)  OR ( Rs("PRIVATE") = 1 and sEName = membername ) and founduser)  Then

									select case dtoday
									case date()
								      Write_TD "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</a>", "Today"
									case cdate(ddate)
								      Write_TD "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</a>", "Selected"
									case else
								      Write_TD "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</a>", "HL"
									end select

									bEvents = True
								    ElseIf dToday < strToDate(Rs("Start_Date")) Then
											Exit Do
										Else	
									    Rs.MoveNext
										End If
								  Loop
									Rs.MoveFirst
								End If
								
								If bEvents = False Then
									select case dtoday
									case date()
									Write_TD "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</a>", "TODAY"
									case cdate(ddate)
									Write_TD "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</a>", "Selected"
									case else
									Write_TD "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</a>", "SOME"
									end select
								End If
							End If
							intPrintDay = intPrintDay + 1
						End If
					Next
					Response.Write "	</tr>" & vbCrLf				
				Loop
				Rs.Close
				set rs = nothing
				Response.Write "</table></td></tr></table></td></tr>"
end function

function emitEvents()
dim  iEvent_Private
if smode = "edit" then
			sSQL = "SELECT start_date, end_date, event_title, event_details, PRIVATE FROM EVENTS WHERE Event_ID = " & clng(Request.QueryString("Event_ID"))
			set rs = server.CreateObject("adodb.recordset")
			Rs.MaxRecords = 1
			Rs.Open sSQL,conn,1,1
	
			If Not Rs.eof Then
				sStart_Date = strToDate(Rs("Start_Date"))
				sEnd_Date = strToDate(Rs("End_Date"))
				sEvent_Title = Rs("Event_Title")
				sEvent_Details = Rs("Event_Details")
				iEvent_Private = RS("PRIVATE")
			end if
			Response.Write "编辑："&sEvent_Title
			Response.Write "</td></tr>"
			Response.Write "<tr><td class=tablebody2>"

			if founduser then
			Show_Form()
			else
			Response.Write "<font color="&Forum_body(8)&">你要成为 " & Forum_info(0) & " 会员才能修改</font><br>"
			Response.Write "<font color="&Forum_body(8)&">请注册为会员, <a href=Reg.asp>注册</a> here.</font>"
			end if
			rs.close
			set rs = nothing
			exit function
end if

if smode = "add" then
		Response.Write "新增我的日记 " & sEvent_Title
		Response.Write "</td></tr>"
		Response.Write "<tr><td class=tablebody2>"

		if founduser then
			Show_Form()
		else
			Response.Write "<font color="&Forum_body(8)&">您必须成为 " & Forum_info(0) & "　的会员才能进行编辑</font><br>"
			Response.Write "<font color="&Forum_body(8)&" >我要成为会员, <a href=REG.asp>注册</a> 这里.</font>"
		end if
		exit function
end if

If IsEmpty(dDate) OR NOT IsDate(dDate) Then
			strMonthName = MonthName(intThisMonth)
			intLastDay = GetLastDay(intThisMonth, intThisYear)
			dFirstDay = intThisMonth & "/1/" & intThisYear
			dLastDay 	= intThisMonth & "/" & intLastDay & "/" & intThisYear

			Response.Write  strMonthName&"日记列表"
			
			sSQL = 	"SELECT * FROM EVENTS  WHERE " & _
							"(Start_Date >='" & DateToStr(dFirstDay) & "' AND Start_Date <= '" & DateToStr(dLastDay) & "') " & _
							"OR " & _
							"(End_Date >='" & DateToStr(dFirstDay) & "' AND End_Date <= '" & DateToStr(dLastDay) & "') " & _
							"OR " & _
							"(Start_Date < '" & DateToStr(dFirstDay) & "' AND End_Date > '" & DateToStr(dLastDay) & "' )"  & _
							"ORDER BY Start_Date"
Else
			Response.Write FormatDateTime(dDate, 1) & "的日记"		
			sSQL = 	"SELECT * FROM EVENTS  " & _
					"WHERE Start_Date <= '" & DateToStr(dDate) & "' AND End_Date >= '" & DateToStr(dDate) & "' ORDER BY Event_ID desc"

End If
Response.Write "</td></tr><tr><td class=tablebody1>"
		if strCMessage <> "" then Response.Write strCMessage
		if sTitle <> "" then Response.Write sTitle
Response.Write "<br></td></tr><tr><td class=tablebody1>"

		dim strMessage
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Rs.Open sSQL,conn,1,1	
		if rs.eof or rs.bof then
			if dDate = "" then
				Response.Write "这个月没有任何日记。"
			else
				Response.Write "今天没有任何日记。"
			end if
		else
		Do While NOT Rs.eof
		sEName = lcase(rs("adduser"))
		if (rs("PRIVATE") <> 1) or (rs("PRIVATE") = 1 and sEName = lcase(membername)) and founduser then%>
		<table  CELLSPACING="1"  CELLPADDING="3" align="center" class=tableBorder2 style="width:98%;">
		<tr>
		<td align="center"  class=tablebody2>
		<%
		Response.Write "<b>"&Trim(Rs("Event_Title"))&"</b>"
		strMessage = Rs("Event_Details")
		%></td>
		<%If (founduser and (lcase(RS("adduser")) = lcase(membername))) or master Then %>
			<td align="center" class=tablebody2 width=10% >
				<a href="z_events.asp?mode=edit&Event_ID=<%= Rs("Event_ID")%>"><img src="<%=Forum_info(7)%>ADDON.gif" BORDER="0" ></a>
				<a href="z_events.asp?mode=delid&Event_ID=<%=Rs("Event_ID")%>&username=<%= HTMLEncode(RS("adduser"))%>&date=<%= HTMLEncode(Request.QueryString("date"))%>"><img src="<%=Forum_info(7)%>delete.gif" BORDER="0" ></a>
			</td>
		<%End If%>
			</tr>
			<tr>
					<td height="30" class=tablebody1 <%If master or lcase(membername) = lcase(RS("adduser")) Then Response.Write "colspan=2" %>>
					<%
											If Rs("Start_Date") <> Rs("End_Date") Then
													Response.Write "日记写于：" & strToDate(Rs("Start_Date")) & vbCrLf
													Response.Write "<br>终止於于：" & strToDate(Rs("End_Date")) & vbCrLf
													Response.Write ""
											else
													Response.Write "<font FACE='宋体, Arial' COLOR='Gray'>日记写于：" & strToDate(Rs("Start_Date")) & vbCrLf
																									
											End If
											Response.Write "　　日记作者：" & RS("adduser")
  											Response.Write "　　天气：" & RS("weather") &"</font><P>"
												%>
				<%=dvbcode(rs("event_details"),UserGroupID,3)%><a href="#top"><img src="<%=Forum_info(7)%>gotop.gif" border="0" align="right" alt="回到页首"></a>
			</td>
			</tr>
			</table>
			<br>
			<% end if
			Rs.MoveNext
			Loop
		End If
		rs.close
		set rs = nothing
end function

function emitupcomingevents
	dim rs,sETitle
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	strSql = "SELECT top 10 event_id,start_date, event_title, adduser,PRIVATE FROM EVENTS  WHERE start_date >= '" & DateToStr(datToday) & "' and start_date < '" & DateToStr(DateAdd("d",30,datToday)) & "' Order by start_date, event_id desc"
	rs.Open strSql,conn,1,1
	do until rs.eof
	sEName = lcase(rs("adduser"))
		if (rs("PRIVATE") <> 1) or (rs("PRIVATE") = 1 and sEName = lcase(membername)) and founduser then
               sETitle=rs("event_title")
		if len(sETitle) > 14 then
			sETitle = mid(sETitle,1,15)
		end if
			Response.Write "<img src=pic/multipage.gif border=0><a href=javascript:showlist('showevents.asp?id=" & rs("event_id") & "',500,500)>" & sETitle & "</a><font color=gray>　－"&sEName&"</font><br>"
		end if
		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
end function

function emitpastEvents
	dim rs,sETitle
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	strSql = "SELECT top 10 event_id,start_date, event_title,adduser,PRIVATE FROM EVENTS  WHERE start_date < '" & DateToStr(datToday) & "' and start_date > '" & DateToStr(DateAdd("d",-30,datToday)) & "' Order by start_date desc"
	rs.Open strSql,conn,1,1
	do until rs.eof
	sEName = lcase(rs("adduser"))
	if (rs("PRIVATE") <> 1) or (rs("PRIVATE") = 1 and sEName = lcase(membername)) and founduser then
	sETitle=rs("event_title")
		if len(sETitle) > 14 then
			sETitle = mid(sETitle,1,15)
		end if
			Response.Write "<img src=pic/multipage.gif border=0><a href=javascript:showlist('showevents.asp?id=" & rs("event_id") & "',500,500)>" & sETitle & "</a><font color=gray>　－"&sEName&"</font><br>"
		end if
		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
end function

Function emitmonths2()
		strMonthName = MonthName(intThisMonth)
		datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
		intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
		intLastDay = GetLastDay(intThisMonth, intThisYear)

		intPrevMonth = intThisMonth - 1
		If intPrevMonth = 0 Then
			intPrevMonth = 12
			intPrevYear = intThisYear - 1
		Else
			intPrevYear = intThisYear	
		End If

		intNextMonth = intThisMonth + 1
		If intNextMonth > 12 Then
			intNextMonth = 1
			intNextYear = intThisYear + 1
		Else
			intNextYear = intThisYear
		End If

		if intThisMonth = "" then
			intLastMonth = DatePart( "m", DateAdd( "m", -1, datToday))
		else
		   	if intThisMonth = 1 then
		   		intLastMonth = 12
			else
				intLastMonth = intThisMonth - 1
			end if
		end if	
		if intThisYear = "" then
			intPrevYear = DatePart( "yyyy", DateAdd( "m", -1, datToday))
		else
			if intThisMonth = 1 then
		  		intPrevYear = intThisYear - 1
			else
		  		intPrevYear = intThisYear
			end if
		end if

		LastMonthDate = GetLastDay(intLastMonth, intPrevYear) - intFirstWeekDay + 2
		NextMonthDate = 1
		intPrintDay = 1
		dFirstDay = intThisMonth & "/1/" & intThisYear
		dLastDay = intThisMonth & "/" & intLastDay & "/" & intThisYear

	dim showbirthday
	showbirthday=false
	if request("showmode")="birthday" then showbirthday=true 
	if showbirthday=false then
		sSQL = 	"SELECT event_id, start_date, end_date, event_title FROM EVENTS  WHERE " & _
						"(Start_Date >='" & DateTostr(dFirstDay) & "' AND Start_Date <= '" & DateTostr(dLastDay) & "') " & _
						"OR " & _
						"(End_Date >='" & DateTostr(dFirstDay) & "' AND End_Date <= '" & DateTostr(dLastDay) & "') " & _
						"OR " & _
						"(Start_Date < '" & DateTostr(dFirstDay) & "' AND End_Date > '" & DateTostr(dLastDay) & "' )"  & _
						"ORDER BY Start_Date"
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open sSQL,conn,1,1
	end if
%>
	<table WIDTH="99%" ALIGN="CENTER" BORDER="1" CELLSPACING="1" CELLPADDING="4" BORDERCOLOR="Gray">
			<tr>
				<td WIDTH="10%" HEIGHT="10" ALIGN="center" VALIGN="MIDDLE" class=tablebody1><a href="<% =ScriptName%>?month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>"><img src="<%=Forum_info(7)%>prev.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="上一个月"></a></td>
				<td width="80%" colspan="5" align="center" valign="middle" class=tablebody2>
					<%=intThisYear%>年 <%=strMonthName%> &nbsp;
				</td>
				<td WIDTH="10%" HEIGHT="10" ALIGN="center" VALIGN="MIDDLE" class=tablebody1><a href="<%=ScriptName%>?month=<%=IntNextMonth%>&amp;year=<%=IntNextYear%>"><img src="<%=Forum_info(7)%>next.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="下一个月"></a></td>
			</tr>
		  <tr>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font color=<%=Forum_body(8)%>>星期日</font></td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle">星期一</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle">星期二</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle">星期三</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle">星期四</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle">星期五</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle">星期六</td>
		  </tr>
		  <%
				EndRows = False
				Response.Write vbCrLf
				
			 	Do While EndRows = False
					Response.Write "	<tr>" & vbCrLf
					For intLoopDay = cSUN To cSAT
						If intFirstWeekDay > cSUN Then
							Write_TD3 LastMonthDate, "NON2"
							LastMonthDate = LastMonthDate + 1
							intFirstWeekDay = intFirstWeekDay - 1
						Else
							If intPrintDay > intLastDay Then
								Write_TD3 NextMonthDate, "NON2"
								NextMonthDate = NextMonthDate + 1
								EndRows = True
							Else
								If intPrintDay = intLastDay Then
									EndRows = True
								End If
								
							dToday = CDate(intThisMonth & "/" & intPrintDay & "/" & intThisYear)
							if showbirthday=false then
									If NOT Rs.eof Then
									bEvents = False
									Do While NOT Rs.eof AND bEvents = False
								    If dToday >= strToDate(Rs("Start_Date")) AND dToday <= strToDate(Rs("End_Date"))  then
										if dtoday = datToday then
										Write_TD2 "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'  > " & intPrintDay & "</a>", "TODAY", dToday
										else
										Write_TD2 "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT' > " & intPrintDay & "</a>", "HL", dToday
										end if
									bEvents = True
									ElseIf dToday < strToDate(Rs("Start_Date")) Then
											Exit Do
									Else	
									    Rs.MoveNext
									End If
									Loop
									Rs.MoveFirst

									End If
									If bEvents = False Then
									if dtoday = datToday then
									Write_TD3 "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</a>	", "TODAY"
									else
									Write_TD3 "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</a>", "SOME"
									end if
									End If
								else
									if dtoday = datToday then
									showbday "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</a>	", "TODAY",dToday
									else
									showbday "<a href=z_events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</a>", "SOME",dToday
									end if
								End If
							End if

							intPrintDay = intPrintDay + 1
						End If
					Next
					Response.Write "	</tr>" & vbCrLf	
				Loop
				if showbirthday=false then
					Rs.Close
					set rs = nothing
				end if
			%>
			<tr>
				<td WIDTH="10%" HEIGHT="10" ALIGN="center" VALIGN="MIDDLE" class=tablebody1><a href="<% =ScriptName%>?month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>"><img src="<%=Forum_info(7)%>prev.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="上一个月"></a></td>
				<td width="80%" colspan="5" align="center" valign="middle" class=tablebody2>&nbsp;
				</td>
				<td WIDTH="10%" HEIGHT="10" ALIGN="center" VALIGN="MIDDLE" class=tablebody1><a href="<%=ScriptName%>?month=<%=IntNextMonth%>&amp;year=<%=IntNextYear%>"><img src="<%=Forum_info(7)%>next.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="下一个月"></a></td>
			</tr>
		</table>
<%
end function

sub mondate(CurrentMonth,stype)
%><!-- 年历开始 -->
	<table WIDTH="99%" Height="100%" ALIGN="CENTER" BORDER="1" CELLSPACING="1" CELLPADDING="4" class=tablebody1 BORDERCOLOR="Gray">
			<%
			dim DayLoop
			dim CurrentMonthName, CurrentYear, FirstDayDate,FirstDay,CurrentDay
			dim DayCounter,CorrectMonth
		if intThisYear = "" then
			CurrentYear = Year(datToday)
		else
			CurrentYear =intThisYear
		end if
			CurrentMonthName = MonthName(CurrentMonth)
			FirstDayDate = DateSerial(CurrentYear, CurrentMonth, 1)
			FirstDay = WeekDay(FirstDayDate, 0)
			CurrentDay = FirstDayDate
			%>
		<tr><th height=23>
		<a href="?month=<%=CurrentMonth%>&year=<%=CurrentYear%>" ><font color=<%=Forum_body(8)%> ><b><%= CurrentMonthName %></b></font></a>
		</th></tr>
		<tr><td class=tablebody1>
			<TABLE COLS="7"  width=100% height=100% CELLPADDING="0" CELLSPACING="2">
			 <tr>
			<% For DayLoop = 1 to 7%>
			<Td ALIGN="CENTER" class=tablebody2 style="border:1px solid #FFB51A">
			<%=Replace(WeekDayName(DayLoop, false,0 ),"星期","")%>
				</Td>
			<% Next%>
			 </Tr>
			<TR>
			<%
			If FirstDay <> 1 Then
			%><TD Colspan="<%=FirstDay -1%>">
			<%
			End if
			
			DayCounter = FirstDay
			CorrectMonth = True
			Do While CorrectMonth = True
			If CurrentDay = datToday Then%>
			<TD align="Center" class=tablebody2 style="border:1px solid #FFB51A; font:12px verdana,宋体;color:blue; padding-left:1px;">
			<% Else %>
			<TD align="Center" class=tablebody1 style="border:1px solid #FFF3C4">
			<% End if%>
			<A Href="z_events.asp?date=<%= Server.URLEncode(CurrentDay)%>&month=<%=CurrentMonth%>&year=<%=CurrentYear%>">
			<%
			if cint(stype)=0 then
			response.write Day(CurrentDay)
					response.write "<br>"
			else
					response.write Day(CurrentDay) 
			end if %>
			</A><br></TD>
			<%
			DayCounter = DayCounter + 1
			If DayCounter > 7 then
				 DayCounter = 1 %>
				 </TR><TR>
			<%
			End if
			CurrentDay = DateAdd("d", 1, CurrentDay)

			If Month(CurrentDay) <> CurrentMonth then
				 CorrectMonth = False
			End if
			Loop
			%>
			<% IF DayCounter <> 1 Then%>
				 <TD Colspan="<%=8-DayCounter%>"> </TD>
			<% End if%>
			</TR>
			<tr><td colspan=7></td>
			</tr>
			</TABLE>
			</td>
		</tr>
		</table>
		<!-- 年历结束 -->
		<%
end sub

function doublenum(fNum)
	if fNum>9 then
		doublenum=fNum
	else
		doublenum="0"&fNum
	end if
end function

function StrToDate(strDateTime)
    dim bstrDateTime
	bstrDateTime=Mid(strDateTime,5,2) & "/" & Mid(strDateTime,7,2) & "/" & Mid(strDateTime,1,4) & " " & Mid(strDateTime,9,2) & ":" & Mid(strDateTime,11,2) & ":" & Mid(strDateTime,13,2)

	if IsDate(bstrDateTime) then
		StrToDate = CDate(bstrDateTime)
	else
		StrToDate = strForumTimeAdjust
	end if
end function

function DateToStr(dtDateTime)
	DateToStr = Year(dtDateTime) & doublenum(Month(dtdateTime)) & doublenum(Day(dtdateTime)) & doublenum(Hour(dtdateTime)) & doublenum(Minute(dtdateTime)) & doublenum(Second(dtdateTime))
end function

call activeonline()
call footer()
%>
