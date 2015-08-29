<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
'=========================================================
'论坛日记列表
' File: userevents.asp
' Version:5.0 Final
' Date:  2002-10-20
' Script Written by sunwin
'=========================================================

stats="浏览日记"
dim username
%>
<html><head>
<meta NAME=GENERATOR Content="Microsoft FrontPage 4.0" CHARSET=GB2312>
<title><%= Forum_info(0)%>--<%=stats%></title>
<!-- #include file="inc/Forum_css.asp" -->
</head>
<body <%=Forum_body(11)%>>
<%
	dim paperid
	dim usersign
    dim    setitle
    dim    sStart_Date
    dim    intPrivateEvent,isPrivateEvent
    dim    sEnd_Date
    dim    sEvent_Title
	dim    sEvent_Details
    dim    EVENT_ID
    dim    Weather
    dim    adduser
	dim	  abgcolor
	if request("id")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
	elseif not isInteger(request("id")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
	else
		paperID=clng(request("id"))
	end if
	if founderr then
		call dvbbs_error()
		response.end
	end if

	if request("orders")="" then
       call main()
	elseif trim(request("orders"))="edit" then
       call edit()
	elseif trim(request("orders"))="save" then
       call save()
	else
	call main()
	end if

sub main()
	if founderr then
		call dvbbs_error()
		exit sub
	else
		set rs=server.createobject("adodb.recordset")
		sql="select * from events where EVENT_ID="&paperid
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>没有找到相关信息。"
		call dvbbs_error()
		exit sub
		else
		if cint(rs("private"))=1  then
			if master OR trim(membername)=trim(rs("adduser")) then
				founderr=false
			else
				founderr=true
			end if
			if founderr=true then 
			Errmsg=Errmsg+"<br>"+"<li>私人日记,暂时保密。"
			call dvbbs_error()
			exit sub
			end if
		end if
        setitle=trim(rs("EVENT_TITLE"))
        if len(sETitle) > 25 then
			sETitle = mid(sETitle,1,26)&"..."
		end if
%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="width:98%">
    <TR> 
      <th>
      <%=htmlencode(setitle)%>
	</th>
	</tr>
        <tr>
        <td  class=tablebody1></td>
        </tr>
    <tr>
      <td  class=tablebody1 valign="top">　
        <table border="0" width="90%" cellspacing="1" cellpadding="0" align=center style="word-break:break-all;"  >
		<tr>
        <td width="100%" class=tablebody1 colspan="2"><font color=gray >发表者：<a href=dispuser.asp?name=<%=htmlencode(rs("adduser"))%> target=_blank><%=htmlencode(rs("adduser"))%></a>　发表时间：<%=strToDate(rs("START_DATE"))%>　天气：<%=htmlencode(rs("weather"))%></font><br>
		<img border="0" src="pic/bt.gif" >
		</td>
          </tr>
          <tr>
            <td width="100%" colspan="2" background="pic\line.gif"><br><br>
			<font class=font1><%=dvbcode(rs("event_details"),UserGroupID,3)%></font></td>
          </tr>
        </table>
      </td>
    </tr>
<% if founduser and htmlencode(membername)=htmlencode(rs("adduser")) then%>
<tr>
      <td  class=tablebody1 align=center><div align=right><a href=?orders=edit&id=<%=rs("event_id")%> title="编辑日记">编辑</a></td>
    </tr>
<% end if %>
    <tr id=tdbg>
      <th align=center><a href=# onclick="window.close();">『 关闭窗口 』</a></th>
    </tr>
  </table>
<%
		end if
		rs.close
		set rs=nothing
	end if

end sub

sub edit()
stats="编辑日记"
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
	call dvbbs_error()
	exit sub
end if
		set rs=server.createobject("adodb.recordset")
		sql="select * from events where EVENT_ID="&paperid
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>没有找到相关信息。"
		call dvbbs_error()
		exit sub
		else
            
		sStart_Date=strToDate(rs("START_DATE"))
		intPrivateEvent=RS("PRIVATE")
		sEnd_Date=strToDate(Rs("End_Date"))
		sEvent_Title = trim(Rs("Event_Title"))
		sEvent_Details = Rs("Event_Details")
		EVENT_ID=rs("EVENT_ID")
		adduser=trim(rs("adduser"))

	if trim(membername)<>trim(rs("adduser")) then
  	errmsg=errmsg+"<br>"+"<li>请您不要修改别人的日记。"
	founderr=true
	call dvbbs_error()
	exit sub
	end if
%><br>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
<tr id=tdbg>
<th colspan=2 align=center>编辑日记</th>
</tr>
<form NAME="EDIT" ACTION="?orders=save" METHOD="POST">
<tr>
<td class=tablebody1 VALIGN="TOP" ALIGN="right" WIDTH="25%">开始日期：</td>
<td class=tablebody1 VALIGN="TOP" ALIGN="LEFT" WIDTH="75%" >
<input TYPE="TEXT" SIZE="12" MAXLENGTH="12" NAME="START_DATE" VALUE="<%if HTMLEncode(sStart_Date) <> "" then Response.write HTMLEncode(sStart_Date) else Response.Write(Request.QueryString("date")) end if%>">
<input type="Checkbox" name="isPrivateEvent" value="1" <% if intPrivateEvent=1 then response.write "checked" %>>私人事项</td>
		</tr>
		<tr>
			<td class=tablebody1 VALIGN="TOP" ALIGN="RIGHT" WIDTH="30%">结束日期：</td>
			<td class=tablebody1 VALIGN="TOP" ALIGN="LEFT" WIDTH="70%">
                        <input TYPE="TEXT" SIZE="12" MAXLENGTH="12" NAME="END_DATE" VALUE="<%if HTMLEncode(sEnd_Date) <> "" then Response.write HTMLEncode(sEnd_Date) else Response.Write(Request.QueryString("date")) end if%>">
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
			<td class=tablebody2 VALIGN="TOP" ALIGN="RIGHT">事项标题：</td>
			<td class=tablebody1 VALIGN="TOP" ALIGN="LEFT">
<%	if sEvent_Title<>"" then
		sEvent_Title=replace(sEvent_Title,"<br>",chr(13))
		sEvent_Title=replace( sEvent_Title,"&nbsp;"," ")
     	end if
%>
 <input TYPE="TEXT" SIZE="30" MAXLENGTH="100" NAME="EVENT_TITLE" VALUE="<%=sEvent_Title%>"></td>
		</tr>
		<tr>
			<td class=tablebody2 VALIGN="TOP" ALIGN="RIGHT">事项内容：</td>
			<td class=tablebody1 VALIGN="TOP" ALIGN="LEFT">
                <textarea COLS="60" ROWS="15" NAME="EVENT_DETAILS" WRAP="PHYSICAL">
<%	if sEvent_Details<>"" then
            sEvent_Details=replace(sEvent_Details,"<br>",chr(13))
            response.write sEvent_Details+chr(13)
	end if
%>
</textarea></td>
		</tr>
		<tr >
			<td class=tablebody1>&nbsp;<% =Request.QueryString("ID") %></td>
			<td class=tablebody2 align=center>
			<input TYPE="hidden" NAME="id" VALUE="<% =Request.QueryString("ID") %>">
                        <input TYPE="hidden" NAME="adduser" VALUE="<% =adduser %>">
			<input TYPE="SUBMIT" VALUE="提交编辑" id="SUBMIT1" name="SUBMIT1">
			</td>
		</tr>
			</form>
	</table>
<% 
    end if
end sub

sub save()
stats="编辑日记完成"
adduser=Checkstr(trim(request.form("adduser")))
Weather=Trim(request.form("Weather"))
sEvent_Title=Trim(request.form("EVENT_TITLE"))
sEvent_Details=trim(Request.Form("EVENT_DETAILS"))
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
	call dvbbs_error()
	exit sub
end if

if  adduser<>membername then 
      Errmsg=Errmsg+"<br>"+"<li>你是不是作者,不能进行修。"
      call dvbbs_error()
      exit sub
end if
if sEvent_Title="" or sEvent_Details="" then
      Errmsg=Errmsg+"<br>"+"<li>请填写主题或内容!"
      call dvbbs_error()
      exit sub
end if
if Request.Form("isPrivateEvent") <> "1" then
		intPrivateEvent = 0
	else
		intPrivateEvent = 1
end if
	set rs=server.createobject("adodb.recordset")
	sql="select EVENT_TITLE,EVENT_DETAILS,PRIVATE,weather from [events] where  EVENT_ID="&paperid
	rs.open sql,conn,1,3
        if NOT(rs.eof and rs.bof) then
		rs(0) =sEvent_Title
		rs(1) =sEvent_Details
        rs(2) =intPrivateEvent
        rs(3) =Weather
        rs.Update
        Errmsg="更新成功，请继续其他操作。"
        call finish()
	else
        Errmsg="更新失败!"
        call finish()
	end if
	rs.Close
	set rs=nothing

end sub

sub finish()

%>
  <table cellpadding=6 cellspacing=1 align=center class=tableborder1>
<tr >
      <th align="center">更新信息
</th></tr>       
<tr>
      <td  height="50" class=tablebody1 align="center">
<%=Errmsg%>
</td>
    </tr>
<tr >
      <th  align="center">
<a href=javascript:history.go(-1)><< 返回上一页</a> ||
<a href=# onclick="window.close();"> 关闭窗口 >></a>
</th></tr>
  </table>
<br>

<%
end sub 

function StrToDate(strDateTime)
        dim bstrDateTime
	bstrDateTime=Mid(strDateTime,5,2) & "/" & Mid(strDateTime,7,2) & "/" & Mid(strDateTime,1,4) & " " & Mid(strDateTime,9,2) & ":" & Mid(strDateTime,11,2) & ":" & Mid(strDateTime,13,2)

	if IsDate(bstrDateTime) then
		StrToDate = CDate(bstrDateTime)
	else
		StrToDate = strForumTimeAdjust
	end if
end function

%>
<%call footer()%>

