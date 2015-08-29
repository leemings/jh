<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<%
'=========================================================
'论坛日记列表
' File: userevents.asp
' Version:6.0
' Date:  2003-3-8
' Script Written by FSsunwin
'=========================================================

     dim UserName
     dim currentpage,page_count,Pcount
     dim totalrec,endpage
     dim EVNum,setitle,adddate
     dim rs1,sql1
     dim stype

	UserName=trim(checkStr(request("evname")))
	if  username="" then
		stats="论坛日记列表"
	else	
		stats="查看"&htmlencode(username)&"的个人日记"
	end if

	if founderr then
		call nav()
		call head_var(2,0,"","")
		call error()
	else	
	call nav()
	call head_var(2,0,"","")
		if request("action")="dellog" then
		call delevents()
		else
		call main()
		end if
        if founderr then call error()
        end if
        call footer()

sub main()
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
function showlist(url, width, height){
	var Win = window.open(url,"showlist",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
//-->
</script>
<table cellspacing=1 cellpadding=3  align=center>
<tr>
<form name="form1" method="post" action="userevents.asp"> 
<td align="left" >
<input type="text" name="evname" size="20">
<input type="submit" name="Submit"  value="用户搜索">                
</td></form>
<td  height="21" align="RIGHT">
<% if  founduser then %>
『<a href="z_events.asp?mode=add&date=<%=Response.Write(Date())%>">新增日记</a>』
<% end if %>
</td>
</tr>
</table>
<table class=tableborder1   cellspacing=1 cellpadding=3  align=center>
<tr>
<th width="5%"  align="center" >状态</th>
<th width="20%"  align="center" >发表人</th>
<th width="50%"  align="center" >主题</th>
<th width="5%"  align="center" >天气</th>
<th width="20%"  align="center" id=tabletitlelink>
<% if founduser then %>
<a href="userevents.asp?action=batch&evname=<%=membername%>&page=<%=request("page")%>">时间</a>
<% else %>
时间
<% end if %>
<% if master then %>
 | <a href="userevents.asp?action=batch&page=<%=request("page")%>">管理</a>
<% end if %>
</th>
</tr>
<% if founduser then %>
<tr><td class=tablebody1 colspan=5 align=center>点击时间,进行日记管理操作!
</td>
</tr>
<% end if
    currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
		set rs=server.createobject("adodb.recordset")
		EVNum=0
		if username<>"" then
		rs1=conn.execute("Select Count(EVENT_ID) From [events] where adduser = '"&checkstr(trim(username))&"'")
		else
		rs1=conn.execute("Select Count(EVENT_ID) From [events] ")
		end if
		EVNum=rs1(0)
		set rs1=nothing
		if isnull(EVNum) then EVNum=0
		if username<>"" then
			if currentpage=1 then
			sql="select top "&Forum_Setting(11)&"  event_id,DATE_ADDED,START_DATE,END_DATE,EVENT_TITLE,PRIVATE,weather,adduser from [events] where adduser = '"&checkstr(trim(username))&"' order by EVENT_ID desc"
			else
			sql="select top "&Forum_Setting(11)&"  event_id,DATE_ADDED,START_DATE,END_DATE,EVENT_TITLE,PRIVATE,weather,adduser from [events] where EVENT_ID<all(select top "&(currentPage-1)*Forum_Setting(11)&" EVENT_ID from [events] where adduser = '"&checkstr(trim(username))&"'  order by EVENT_ID desc) and adduser = '"&username&"' order by EVENT_ID desc"
			end if
		else
			if currentpage=1 then
			sql="select top "&Forum_Setting(11)&"  event_id,DATE_ADDED,START_DATE,END_DATE,EVENT_TITLE,PRIVATE,weather,adduser from [events] order by EVENT_ID desc"
			else
			sql="select top "&Forum_Setting(11)&"  event_id,DATE_ADDED,START_DATE,END_DATE,EVENT_TITLE,PRIVATE,weather,adduser from [events] where EVENT_ID<all(select top "&(currentPage-1)*Forum_Setting(11)&" EVENT_ID from [events] order by EVENT_ID desc)  order by EVENT_ID desc"
			end if
		end if
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
%>
<tr> 
<td colspan=5 class=tablebody1>　　还没有任何日记数据或你还未登陆.</td>
</tr>
<%
		else
		while (not rs.eof)
		setitle=rs("EVENT_TITLE")
		if len(sETitle) > 25 then
		sETitle = mid(sETitle,1,26)&"..."
		end if
		adddate=strToDate(rs("DATE_ADDED"))
%>
<form action="userevents.asp?action=dellog&evname=<%=username%>" method=post name=even>
<tr>
<td  height="21" align="center" class=tablebody1><% if rs("private")=0 then %><img src=pic/on.gif alt="公开，欢迎浏览！"><% else %><img src=pic/offlock.gif alt="保密，只有作者可以观看！"><%end if%></td>
<td  height="1"  class=tablebody2 align="center"><a href=dispuser.asp?name=<%=rs("adduser")%> target=_blank><%=rs("adduser")%></a></td>
<td  height="1" align="left" class=tablebody1>
<%
	if master or rs("private")=0 then
	response.write "<a href=javascript:showlist('showevents.asp?id="&rs("event_id")&"',500,500)>"
	elseif founduser and trim(membername)=trim(rs("adduser")) then 
	response.write"<a href=javascript:showlist('showevents.asp?id="&rs("event_id")&"',500,500)>"
end if
response.write setitle
%>
</a></td>
<td  height="1" class=tablebody1 align="center"><%=rs("weather")%></td>
<td  height="1" class=tablebody2 align="left">
<%if request("action")="batch" then%><input type="checkbox" name="delid" value="<%=rs("event_id")%>"><%end if%>
<%=adddate%>
</td>
</tr>
<%
		rs.movenext
		wend
	end if
	if request("action")="batch" then
	response.write "<tr><td  colspan=5 align=right class=tablebody2>请选择要删除的事件，<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">全选 <input type=submit name=Submit value=执行  onclick=""{if(confirm('您确定执行的操作吗?')){this.document.even.submit();return true;}return false;}""></td></tr>"
	end if
%></form>
</table>
<%

		Pcount=int(EVNum/Forum_Setting(11))

		if (EVNum mod Forum_Setting(11))>0 then
			Pcount=Pcount+1
		end if
		response.write "<table  width="&Forum_body(12)&" cellspacing=1 cellpadding=3  align=center >"&_
				"<tr><td valign=middle nowrap >"&_
				"页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
				"　每页显示：<b>"&Forum_Setting(11)&"</b>篇 目前共：<b>"&EVNum&"</b>篇</td>"&_
				"<td valign=middle nowrap ><div align=right><p>分页： "
	
		if currentpage > 4 then
		response.write "<a href=""?evname="&username&"&page=1"">[1]</a> ..."
		end if
		if Pcount>currentpage+3 then
		endpage=currentpage+3
		else
		endpage=Pcount
		end if
		for i=currentpage-3 to endpage
		if not i<1 then
			if i = clng(currentpage) then
	        response.write " <font color="&Forum_body(8)&">["&i&"]</font>"
			else
	        response.write " <a href=""?evname="&username&"&page="&i&""">["&i&"]</a>"
			end if
		end if
		next
		if currentpage+3 < Pcount then 
		response.write "... <a href=""?evname="&username&"&page="&Pcount&""">["&Pcount&"]</a>"
		end if
		response.write "</p></div></font></td></tr></table>"

	rs.close
	set rs=nothing
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

sub delevents()
	dim delid
	if founduser=false then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请登陆后进行操作。"
	end if
       if membername<>trim(checkStr(request("evname")))  then
            if  not master then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>您不是该日记作者或者系统管理员。" 
            end if           
	end if
	if request("DELid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关日记。"
	else
		DELid=CheckStr(request.Form("DELid"))
	end if
	if founderr then exit sub
	conn.execute("delete from EVENTS where event_id in ("&DELid&")")
	if err.number<>0 then
        	err.clear
		ErrMsg=ErrMsg+"<Br>"+"<li>数据库操作失败，请以后再试:"&err.Description 
  		exit sub
	else
		call success()
	end if
end sub
sub success()
%>
    <table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
    <tr align="center"> 
    <th >成功：日记操作</th>
    </tr>
    <tr> 
    <td  class=tablebody1><li>您选择的日记已经删除。<br>
    </td>
    </tr>
    <tr > 
    <th id=tabletitlelink>
    <a href="USEREVENTS.asp"> << 返回日记列表页面</a>
    </th></tr></table>
<%
end sub
%>

