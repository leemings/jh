<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<%
dim stype,ordername
dim currentpage,page_count,Pcount
dim totalrec,endpage
dim searchDateLimit,searchday
dim searchboard
dim LastPost,Lastuser,LastID
dim LastTime,LastUserid,LastRootid,body

stats="热门话题"

stype=request("stype")
if request("page")="" or not isInteger(request("page"))  then
	currentPage=1
else
	currentPage=cint(request("page"))
end if
'搜索多少天内帖子
searchDateLimit=180
if request("SearchDate")="all" OR request("SearchDate")="" then
	searchday=" "
else
	if not isInteger(request("SearchDate"))  then
		Errmsg=Errmsg+"<br>"+"<li>搜索多少天必须是整数。"
		founderr=true
	else
		searchday=" datediff('d',dateandtime,Now()) < "&request("SearchDate")&" and "
	end if
end if
call nav()
if cint(boardid)<1 then
call head_var(2,0,"","hotlist.asp?boardid=0")
searchboard=""
else
call head_var(1,0,boardtype,"hotlist.asp?boardid="&boardid)
searchboard=" boardid="&boardid&" and "
end if

if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛搜索的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
end if
'###################特殊版面修改开始(asilas制作)##################
if boardid>0 then
	if cint(Board_Setting(51))=1 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
		else
			if chkviplogin(membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>VIP会员专用版面</font>，请确认您的属性是否符合。"
			founderr=true
			end if
		end if
	end if
	if cint(Board_Setting(52))<>0 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
		else
			dim sexshow
			if cint(Board_Setting(52))=1 then
				sexshow="女生"
			elseif cint(Board_Setting(52))=2 then
				sexshow="男生"
			end if
			if chksexlogin(cint(Board_Setting(52)),membername)=false then
				Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>"&sexshow&"论坛版面</font>，请确认您的性别是否符合。"
				founderr=true
			end if
		end if
	end if
end if
'####################特殊版面修改结束(asilas制作)#################
if founderr then
	call dvbbs_error()
else
	call search()
	if founderr then call dvbbs_error()
end if
call footer()

sub search()
	sql=""
	set rs=server.createobject("adodb.recordset")
	select case stype
	case 1
		sql="select top 50  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,isvote,isbest,locktopic,Expression,istop from Topic where "&searchboard&" "&searchday&" locktopic<2 and child>"&forum_setting(44)&" ORDER BY child desc"
		ordername="最热门50主题"
	case 2
		sql="select top 50  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,isvote,isbest,locktopic,Expression,istop from Topic where "&searchboard&" "&searchday&" locktopic<2 ORDER BY child desc"
		ordername="最新50帖"
	case else
		sql="select top 50  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,isvote,isbest,locktopic,Expression,istop from Topic where "&searchboard&" "&searchday&" locktopic<2 and child>"&forum_setting(44)&" ORDER BY child desc"
		ordername="最热门50主题"
	end select
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>没有找到您要查询的内容。"
		founderr=true
		exit sub
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		call searchinfo()
		call listPages3()
	end if
	rs.close
	set rs=nothing
	call activeonline()
end sub

sub searchinfo()%>
	<form action="<%=ScriptName%>" method=get><table cellpadding=0 cellspacing=0 border=0 width="<%=Forum_body(12)%>" align=center>
	<tr><td>目前共有<font color=<%=Forum_body(8)%>><%=totalrec%></font>个<%=ordername%>
	</td>
	<td width=200 align=right>
	<input type=hidden name=boardid value="<%= boardid %>">
	<input type=hidden name=stype value="<%=request("stype")%>">
	<select name=SearchDate onchange='javascript:submit()'>
	<option value=all <%if request("SearchDate")="all" then%>selected<%end if%>>查看所有的主题
	<option value=1 <%if request("SearchDate")="1" then%>selected<%end if%>>查看一天内的主题
	<option value=2 <%if request("SearchDate")="2" then%>selected<%end if%>>查看两天内的主题
	<option value=7 <%if request("SearchDate")="7" then%>selected<%end if%>>查看一星期内的主题
	<option value=15 <%if request("SearchDate")="15" then%>selected<%end if%>>查看半个月内的主题
	<option value=30 <%if request("SearchDate")="30" then%>selected<%end if%>>查看一个月内的主题
	<option value=60 <%if request("SearchDate")="60" then%>selected<%end if%>>查看两个月内的主题
	<option value=180 <%if request("SearchDate")="180" then%>selected<%end if%>>查看半年内的主题
	</select>
	</td></tr>
	</table></form>
	<script>
	function loadThreadFollow(t_id,b_id){
		var targetImg =eval("document.all.followImg" + t_id);
		var targetDiv =eval("document.all.follow" + t_id);
	
		if ("object"==typeof(targetImg)){
			if (targetDiv.style.display!='block'){
				targetDiv.style.display="block";
				targetImg.src="<%=Forum_info(7)&Forum_boardpic(16)%>";
				if (targetImg.loaded=="no"){
					document.frames["hiddenframe"].location.replace("loadtree1.asp?boardid="+b_id+"&rootid="+t_id);
				}
			}else{
				targetDiv.style.display="none";
				targetImg.src="<%=Forum_info(7)&Forum_boardpic(15)%>";
			}
		}
	}
	</script>
	<iframe width=0 height=0 src="" id="hiddenframe"></iframe>
	<TABLE cellPadding=3 cellSpacing=1 class=tableborder1 align=center>
	<TR valign=middle>
	<Th height=25 width=32>状态</Th>
	<Th width=*>主 题</Th>
	<Th width=80>作 者</Th>
	<Th width=64>回复/人气</Th>
	<Th width=70>最后更新</Th>
	<Th width=80>回复人</Th>
	</TR>
	<%while (not rs.eof) and (not page_count = rs.PageSize)
		if not isnull(rs("lastpost")) then
			LastPost=split(rs("lastpost"),"$")
			if ubound(LastPost)=6 then
			Lastuser=htmlencode(LastPost(0))
			LastID=LastPost(1)
			LastTime=LastPost(2)
			if not isdate(LastTime) then LastTime=rs("dateandtime")
			body=htmlencode(LastPost(3))
			LastUserid=LastPost(5)
			LastRootid=LastPost(6)
			else
			Lastuser=htmlencode(LastPost(0))
			LastID=rs("topicid")
			LastTime=rs("dateandtime")
			body="..."
			LastUserid=rs("postuserid")
			LastRootid=rs("topicid")
			end if
		else
			Lastuser="------"
			LastID=rs("topicid")
			LastTime=rs("dateandtime")
			body="..."
			LastUserid=rs("postuserid")
			LastRootid=rs("topicid")
		end if
		response.write"<TR><TD align=middle class=tablebody2 width=32 height=27>"
		response.write TopicIcon(rs("BoardID"),rs("topicid"),rs("lastposttime"),rs("istop"),rs("isVote"),rs("LockTopic"),iif(Rs("child")>=Cint(forum_setting(44)),1,0))
		response.write"</TD><TD onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1' class=tablebody1 width=*>"
		'response.write "&nbsp;<img src='"& Forum_info(8) & rs("Expression") &"'>&nbsp;"
		if renzhen(rs("boardid"),membername)=false then
			response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& Rs("topicid") &""">"
			response.write "<a href='dispbbs.asp?boardID="&rs("boardID")&"&ID="& rs("topicid") &"&skin=1'>"
			response.write "<font color=gray>（认证或隐藏论坛帖子，只有认证用户或版主才能查看）</font>"
		elseif VipBoard(rs("boardid"),membername)=false then
			response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& Rs("topicid") &""">"
			response.write "<a href='dispbbs.asp?boardID="&rs("boardID")&"&ID="& rs("topicid") &"&skin=1'>"
			response.write "<font color=gray>（VIP或特殊论坛帖子，只有VIP或特殊会员才能查看）</font>"
		else
			if Rs("child")=0 then
				response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& Rs("topicid") &""">"
			else
				response.write "<img loaded=""no"" src="""& Forum_info(7)&Forum_boardpic(15)&""" id=""followImg"& Rs("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& Rs("topicid") &","& Rs("boardid") &")"" title=展开帖子列表>"
			end if
			response.write "<a href='dispbbs.asp?boardID="&rs("boardID")&"&ID="& rs("topicid") &"&skin=1'>"
			if len(rs("title"))>26 then
				response.write ""&left(reubbcode(htmlencode(replace(rs("title"),chr(10)," ")),true),26)&"..."
			else
				response.write reubbcode(htmlencode(rs("title")),true)
			end if
		end if
		response.write "</a>"
		if rs("isbest")=1 then
			response.write "&nbsp;&nbsp;<img src="""&Forum_info(7)&Forum_statePic(4)&""">"
		end if
		response.write "</TD>"
		response.write "<TD align=middle  class=tablebody2  width=80><a href=""dispuser.asp?id="&rs("postuserid")&""">"&htmlencode(rs("postusername"))&"</a></TD>" 
		response.write "<TD class=tablebody1 width=64 align=center>"
		if rs("isvote")=1 then
			response.write "<FONT color="&Forum_body(8)&"><b>"&rs("votetotal")&"</b></font>  票"
		else
			response.write rs("child") &"/"& rs("hits")
		end if
		response.write "</TD>"%>
		<TD align=center class=tablebody2 width=70><a href='dispbbs.asp?boardid=<%=rs("boardid")%>&id=<%=LastRootID%>&<%=LastID%>'><%=month(LastTime)&"-"&day(LastTime)%>&nbsp;<%=FormatDateTime(LastTime,4)%></a></TD>
		<TD align=center class=tablebody1 width=80><a href="dispuser.asp?id=<%=LastPost(5)%>" target=_blank><%=htmlencode(LastUser)%></a></TD>
		</TR> 
		<%response.write "<tr style=""display:none"" id=""follow"& Rs("topicid") &"""><td colspan=6 id=""followTd"& Rs("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& Rs("topicid") &","&Rs("boardid")&")"">正在读取关于本主题的跟帖，请稍侯……</div></td></tr>"
		page_count = page_count + 1
		rs.movenext
	wend
	response.write "</table>"
end sub

sub listPages3()
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<b>"&Forum_Setting(11)&"</b> 帖子数<b>"&totalrec&"</b></font></td>"&_
			"<td valign=middle nowrap><font color="&Forum_body(13)&"><div align=right><p>分页： "
	call DispPageNum(currentpage,PCount,"""?page=","&stype="&stype&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"""")
	response.write "</p></div></font></td></tr></table>"
end sub%>