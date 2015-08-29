<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
dim currentpage,page_count,Pcount
dim totalrec,endpage
dim stype,pSearch,nSearch,keyword
dim searchboard,ordername
dim searchdate,searchday
dim stable
dim scoreuser,searchscore,searchscoreuser
stats="搜索结果"
stype=request("stype")
pSearch=request("pSearch")
nSearch=request("nSearch")
keyword=trim(checkStr(request("keyword")))
stable=checkstr(request("stable"))
scoreuser=trim(checkStr(request("scoreuser")))
if stable="" then stable=Nowusebbs
if request("page")="" or not isInteger(request("page"))  then
	currentPage=1
else
	currentPage=cint(request("page"))
end if
if stype<3 then
	if keyword="" and scoreuser="" and request("SearchScore")="ALL" then
		Errmsg=Errmsg+"<br>"+"<li>必须输入查询关键字。"
		founderr=true
	end if
	'搜索多少天内帖子
	if request("SearchDate")="ALL" then
		searchday=" "
	else
		if not isInteger(request("SearchDate"))  then
			Errmsg=Errmsg+"<br>"+"<li>搜索多少天必须是整数。"
			founderr=true
		else
			searchday=" datediff('d',b.dateandtime,Now()) < "&request("SearchDate")&" and "
		end if
	end if
	if scoreuser="" then searchscoreuser=" "
	select case request("SearchScore")
	case "ALL"
		SearchScore=" "
	case "NONE"
		SearchScore=" isnull(Score) and "
	case "A5"
		SearchScore=" not isnull(Score) and Score>=5 and "
	case "A4"
		SearchScore=" not isnull(Score) and Score>=4 and "
	case "A3"
		SearchScore=" not isnull(Score) and Score>=3 and "
	case "A2"
		SearchScore=" not isnull(Score) and Score>=2 and "
	case "A1"
		SearchScore=" not isnull(Score) and Score>=1 and "
	case "Z0"
		SearchScore=" not isnull(Score) and Score=0 and "
	case "B1"
		SearchScore=" not isnull(Score) and Score<=-1 and "
	case "B2"
		SearchScore=" not isnull(Score) and Score<=-2 and "
	case "B3"
		SearchScore=" not isnull(Score) and Score<=-3 and "
	case "B4"
		SearchScore=" not isnull(Score) and Score<=-4 and "
	case "B5"
		SearchScore=" not isnull(Score) and Score<=-5 and "
	end select
end if
call nav()
if cint(boardid)<1 then
	call head_var(0,0,"论坛搜索","query.asp?boardid=0")
	searchboard=""
else
	call head_var(0,0,boardtype,"query.asp?boardid="&boardid)
	searchboard=" b.boardid="&boardid&" and "
end if

if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛搜索的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
end if
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
		select case nSearch
		'搜索主题帖子作者
		case 1
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid=0 and b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where b.username='"&keyword&"' and "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid=0 and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="搜索主题作者帖子"
		'搜索回复帖子作者
		case 2
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid>0 and b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where b.username='"&keyword&"' and "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.parentid>0 and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="搜索回复作者帖子"
		'两者都搜索
		case 3
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.username='"&keyword&"' and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="搜索主题和回复作者帖子"
		end select
	case 2
		select case pSearch
		'搜索主题关键字
		case 1
			if keyword="" then
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" (" & translate(keyword,"b.topic") & ") and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			'搜索内容关键字
			ordername="搜索主题"
		case 2
			if keyword="" then
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" (" & translate(keyword,"b.body") & ") and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			'两者都搜索
			ordername="搜索内容"
		case 3
			if keyword="" then
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" b.locktopic<2 ORDER BY b.announceID desc"
			else
				sql="select top 1000  b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" "&searchday&" "&searchscoreuser&" "&searchscore&" (" & translate(keyword,"b.topic") & " or " & translate(keyword,"b.body") & ") and b.locktopic<2 ORDER BY b.announceID desc"
			end if
			ordername="搜索主题和内容"
		end select
	case 3
		if searchboard="" then
			sql="select top 100 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" b.locktopic<2 ORDER BY b.announceID desc"
			ordername="最新100帖"
		else
			sql="select top 30 b.locktopic,b.boardid,b.rootid,b.announceid,b.body,b.Expression,b.topic,b.username,b.dateandtime,b.postuserid,b.ParentID,b.isbest,f.boardtype from "&stable&" b inner join board f on b.boardid=f.boardid where  "&searchboard&" b.locktopic<2 ORDER BY b.announceID desc"
			ordername="最新30帖"
		end if
	end select
	
	if sql="" then
		Errmsg=Errmsg+"<br>"+"<li>请指定查询条件。"
		founderr=true
		exit sub
	end if
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

sub searchinfo()
	%><table cellpadding=0 cellspacing=0 border=0 width="<%=Forum_body(12)%>" align=center>
		<tr>
			<td>查询<%if request("SearchDate")="ALL" then%>所有日期<%else%><%=request("SearchDate")%>天内<%end if%>的帖子，<%=ordername%><%if totalrec>=1000 then%>最多只能查询到<font color=<%=Forum_body(8)%>>1000</font>条纪录，更多的就不显示了<%else%>共查询到<font color=<%=Forum_body(8)%>><%=totalrec%></font>个结果<%end if%></td>
		</tr>
	</table>
	<TABLE cellPadding=3 cellSpacing=1 class=tableborder1 align=center>
		<TR valign=middle>
			<Th height=25 width=32>状态</Th>
			<Th width=*>主 题</Th>
			<Th width=80>论坛</Th>
			<Th width=80>作 者</Th>
			<Th width=70>最后更新</Th>
			<Th width=80>回复人</Th>
		</TR><%
		while (not rs.eof) and (not page_count = rs.PageSize)
			%><TR> 
				<TD align=center class=tablebody2 width=32><%
					dim RsTopic,lock
					lock=0
					if rs(10)=0 then
						set RsTopic=conn.execute("select * from topic where topicid="&rs(2))
						response.write TopicIcon(rs("BoardID"),rs("rootid"),rs("DateAndTime"),rstopic("istop"),rstopic("isVote"),rstopic("LockTopic"),iif(RsTopic("child")>=Cint(forum_setting(44)),1,0))
						Set rsTopic=nothing
					else
						response.write TopicIcon(rs("BoardID"),rs("rootid"),rs("DateAndTime"),0,0,0,0)
					end if
				%></TD>
				<TD onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1'  class=tablebody1 width=*>&nbsp;<a href='dispbbs.asp?boardID=<%=rs(1)%>&replyID=<%=rs(3)%>&ID=<%=rs(2)%>&skin=1' target=_blank><img src='<%=Forum_info(8)&rs(5)%>' border=0 alt="开新窗口浏览此主题"></a>&nbsp;<a href='dispbbs.asp?boardID=<%=rs(1)%>&replyID=<%=rs(3)%>&ID=<%=rs(2)%>&skin=1'><%
					if renzhen(rs(1),membername)=false then
						%><font color=gray>（认证或隐藏论坛帖子，只有认证用户或版主才能查看）</font><%
					elseif VipBoard(rs(1),membername)=false then
						%><font color=gray>（VIP或特殊论坛帖子，只有VIP或特殊会员才能查看）</font><%
					else
						if rs(6)="" then
							if rs("isbest")=1 and Cint(GroupSetting(41))=0 then
								%><font color=gray>（精华帖，禁止预览）</font><%
							elseif rs("isbest")=2 then
								%><font color=gray>（单帖屏蔽，禁止预览）</font><%
							else
								%><%=left(htmlencode(replace(reUBBCode(rs(4),false),chr(10)," ")),26)%>...<%
							end if
						else
							%><%=htmlencode(reUBBCode(rs(6),false))%><%
						end if
					end if
				%></a></TD> 
    		<TD align=center class=tablebody2  width=80><a href="list.asp?boardid=<%=htmlencode(rs("boardid"))%>"><%=htmlencode(rs("boardtype"))%></a></TD>
    		<TD align=center class=tablebody1  width=80><a href="dispuser.asp?id=<%=htmlencode(rs("postuserid"))%>"><%=htmlencode(rs(7))%></a></TD>
    		<TD align=center class=tablebody2 width=70><%=month(rs("dateandtime"))&"-"&day(rs("dateandtime"))%>&nbsp;<%=FormatDateTime(rs("dateandtime"),4)%></TD>
    		<TD align=center class=tablebody1 width=80>------</TD>
			</TR><%
	  	page_count = page_count + 1
			rs.movenext
		wend
	%></table><%
end sub

	sub listPages3()

	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<b>"&Forum_Setting(11)&"</b> 帖子数<b>"&totalrec&"</b></font></td>"&_
			"<td valign=middle nowrap><font color="&Forum_body(13)&"><div align=right><p>分页： "
	call DispPageNum(currentpage,PCount,"""?page=","&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&"&scoreuser="&scoreuser&"&searchscore="&request("SearchScore")&"""")
	response.write "</p></div></font></td></tr></table>"

	end sub

public function translate(sourceStr,fieldStr)
rem 处理逻辑表达式的转化问题
  dim  sourceList
  dim resultStr
  dim i,j
  if instr(sourceStr," ")>0 then 
     dim isOperator
     isOperator = true
     sourceList=split(sourceStr)
     '--------------------------------------------------------
     rem Response.Write "num:" & cstr(ubound(sourceList)) & "<br>"
     for i = 0 to ubound(sourceList)
        rem Response.Write i 
    Select Case ucase(sourceList(i))
    Case "AND","&","和","与"
        resultStr=resultStr & " and "
        isOperator = true
    Case "OR","|","或"
        resultStr=resultStr & " or "
        isOperator = true
    Case "NOT","!","非","！","！"
        resultStr=resultStr & " not "
        isOperator = true
    Case "(","（","（"
        resultStr=resultStr & " ( "
        isOperator = true
    Case ")","）","）"
        resultStr=resultStr & " ) "
        isOperator = true
    Case Else
        if sourceList(i)<>"" then
            if not isOperator then resultStr=resultStr & " and "
            if inStr(sourceList(i),"%") > 0 then
                resultStr=resultStr&" "&fieldStr& " like '" & replace(sourceList(i),"'","''") & "' "
            else
                resultStr=resultStr&" "&fieldStr& " like '%" & replace(sourceList(i),"'","''") & "%' "
            end if
                isOperator=false
        End if    
    End Select
        rem Response.write resultStr+"<br>"
     next 
     translate=resultStr
  else '单条件
     if inStr(sourcestr,"%") > 0 then
         translate=" " & fieldStr & " like '" & replace(sourceStr,"'","''") &"' "
     else
    translate=" " & fieldStr & " like '%" & replace(sourceStr,"'","''") &"%' "
     End if
     rem 前后各加一个空格，免得连sql时忘了加，而出错。
  end if  
end function
%>