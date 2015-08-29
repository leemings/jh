<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
'=========================================================
' File: admin_topiclist.asp
' Version:5.0
' Date: 2002-9-20
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

stats="帖子审核"
Dim currentPage
dim AdminLockTopic
dim p,announceIDRange1,announceIDRange2,tableclass
dim bBoardEmpty
bBoardEmpty=false
AdminLockTopic=false
if (master or superboardmaster or boardmaster) and Cint(GroupSetting(36))=1 then
	AdminLockTopic=true
else
	AdminLockTopic=false
end if
if Cint(GroupSetting(36))=1 and UserGroupID>3 then
	AdminLockTopic=true
end if
if FoundUserPer and Cint(GroupSetting(36))=1 then
	AdminLockTopic=true
elseif FoundUserPer and Cint(GroupSetting(36))=0 then
	AdminLockTopic=false
end if
if not AdminLockTopic then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本版面审核帖子的权限。"
	founderr=true
end if
currentPage=request("page")
if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,BoardDepth,0,0)
	if request("action")="freetopic" then
	call freetopic()
	else
	call main()
	end if
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
dim totalrec,ii,page_count
dim n,pi
dim rs1,sql1
%>
<BR>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>
<form action="admin_topiclist.asp?action=freetopic" method=post name=batch>
<input type=hidden value="<%=boardid%>" name=boardid>
<TR align=middle>
<Th height=25 width=32 id=tabletitlelink><a href="list.asp?boardid=<%=boardid%>&page=<%=request("page")%>&action=batch">选项</a></th>
<Th width=*>主 题</Th>
<Th width=80>作 者</Th>
</TR>
<%
set rs=server.createobject("adodb.recordset")
sql="select rootid from "&NowUseBBS&" where boardid="&boardid&" and ParentID=0 and isAudit=1 order by Announceid desc"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write "<tr><td colSpan=3 width=100% class=tablebody1>暂无审核内容</td></tr>"
	rs.close
	set rs=nothing
	bBoardEmpty = true
else
	bBoardEmpty = false
	totalrec=rs.recordcount

	RS.PageSize=Cint(Forum_Setting(11))
	If currentpage <> "" then
		currentpage =  cint(currentpage)
		if currentpage<1 then  
			currentpage = 1
		end if
	else
		currentpage = 1
	End if 
	if currentpage*Cint(Forum_Setting(11))>totalrec and not((currentpage-1)*Cint(Forum_Setting(11))<totalrec)then currentPage=1
	Rs.AbsolutePage = currentpage
	announceIDRange1=rs(0)
	rs.move Cint(Forum_Setting(11))-1
	if rs.EOF then rs.movelast
	announceIDRange2=rs(0)
	rs.close
	set rs=nothing
end if

if not bBoardEmpty then
sql="select AnnounceID,boardID,UserName,Topic,DateAndTime,RootID,layer,orders,Expression,body,PostUserID,locktopic from "&NowUseBBS&" where BoardID=" & boardID & " and ( rootID >= " & announceIDRange2 &  " and rootID <=" & announceIDRange1 & ") order by RootID desc,Orders"
'response.write sql
set rs=conn.execute(sql)
do while not rs.eof
page_count=page_count+1
if rs("layer")= 1 then
	tableclass="tablebody1"
else
	tableclass="tablebody2"
end if
response.write "<TR align=middle><TD class=tablebody2 width=32 height=27 class="&tableclass&">"
if rs("locktopic")=3 then
response.write "<input type=checkbox name=Announceid value="""&rs("Announceid")&""">"
else
response.write "&nbsp;"
end if
response.write "</TD><TD align=left class=tablebody1 width=* class="&tableclass&">"

if rs("layer")>1 then
	for i=2 to rs("layer")
		response.write "&nbsp;&nbsp;"
	next 
end if

response.write "<img src=face/"&rs("Expression")&"> "
response.write "<a href='dispbbs.asp?action=dispaudit&boardID="&boardID&"&ID="&cstr(rs("RootID"))&"&replyID="&Cstr(rs("announceID"))&"&skin=1'>"

if rs("topic")="" or isnull(rs("topic")) then
	if len(rs("body"))>50 then
		response.write htmlencode(replace(left(rs("body"),50),chr(10),""))
	else
		response.write htmlencode(replace(rs("body"),chr(10),""))
	end if
else
	if len(rs("Topic"))>50 then
		response.write htmlencode(left(rs("Topic"),50))
	else
		response.write htmlencode(rs("Topic"))
	end if
end if
response.write "&nbsp;("&rs("dateandtime")&")</TD>"
response.write "<TD class=tablebody2 width=80 class="&tableclass&"><a href=""dispuser.asp?id="& rs("postuserid") &""" target=_blank>"& htmlencode(rs("username")) &"</a></TD>"

response.write "</TR>"
rs.movenext
loop
end if
set rs=nothing
if totalrec mod Forum_Setting(11)=0 then
   	n= totalrec \ Forum_Setting(11)
else
   	n= totalrec \ Forum_Setting(11)+1
end if
if currentpage-1 mod 10=0 then
	p=(currentpage-1) \ 10
else
	p=(currentpage-1) \ 10
end if
dim pagelist,pagelistbit
%>
<TR align=middle>
<Td height=25 class=tablebody2 colspan=3>&nbsp;请选择要操作的内容：<input name="actiontype" value="1" type=radio checked>通过审核&nbsp;<input name="actiontype" value="2" type=radio>删除帖子&nbsp;<input name=submit value="执行" type=submit onclick="{if(confirm('您确定执行此操作吗?')){this.document.batch.submit();return true;}return false;}"></Td>
</TR>
</table>

<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align="center">
</form>
<form method=post action="admin_topiclist.asp">
<input type=hidden name=selTimeLimit value='<%= request("selTimeLimit") %>'><tr>
<td valign=middle>页次：<b><%= currentPage %></b>/<b><%= n %></b>页 每页<b><%= Forum_Setting(11) %></b> 主题数<b><%= totalrec %></b></td>
<td valign=middle><div align=right >分页：
<%call DispPageNum(currentpage,n,"'?boardid="&boardid&"&page=","&selTimeLimit="&request("selTimeLimit")&"&action="&request("action")&"'")%>
转到:<input type=text name=Page size=3 maxlength=10  value='<%= currentpage %>'><input type=submit value=Go name=submit>
</div></td></tr>
<input type=hidden name=BoardID value='<%= BoardID %>'>
</form></table>

<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr>
<FORM METHOD=POST ACTION="queryresult.asp?boardid=<%=boardid%>&sType=2&SearchDate=30&pSearch=1">
<td width=50% valign=middle nowrap height=40>
快速搜索：<input type=text name="keyword">&nbsp;<input type=submit name=submit value="搜索">
</td>
</FORM>
<td valign=middle nowrap width=50%> <div align=right>
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option selected>跳转论坛至...</option>
<%
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
response.write "<option value=""list.asp?boardid="&rs(0)&""" "
if boardid=rs(0) then
response.write "selected"
end if
response.write ">"
if rs(2)>0 then
for i=1 to rs(2)
response.write "－"
next
end if
response.write rs(1)&"</option>"
rs.MoveNext
loop
set rs=nothing

response.write "</select><div></td></tr></table>"

end sub

sub freetopic()
if request.form("announceid")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
	exit sub
end if
Dim id,trs,ars
dim bbsnum,topicnum,todaynum
dim haveaudit
bbsnum=0
topicnum=0
todaynum=0
for i=1 to request.form("Announceid").count
	ID=replace(request.form("Announceid")(i),"'","")
	if request("actiontype")=2 then
		set rs=conn.execute("select rootid from "&NowUsebbs&" where parentid=0 and Announceid="&id)
		if not (rs.eof and rs.bof) then
			conn.execute("delete from topic where topicid="&rs(0))
			conn.execute("delete from "&NowUseBBS&" where rootid="&rs(0))
		else
			conn.execute("delete from "&NowUseBBS&" where Announceid="&id)
		end if
	elseif cint(request("actiontype"))=1 then
		set rs=conn.execute("select rootid,dateandtime,PostUserID from "&NowUsebbs&" where parentid=0 and Announceid="&id)
		if not (rs.eof and rs.bof) then
			bbsnum=bbsnum+1
			topicnum=topicnum+1
			if datediff("d",rs(1),Now())=0 then todaynum=todaynum+1
			conn.execute("update topic set locktopic=0 where topicid="&rs(0))
			conn.execute("update "&NowUseBBS&" set locktopic=0,isaudit=0 where Announceid="&id)
			conn.execute("update [user] set article=article+1,userWealth=userWealth+"&Forum_user(2)&",UserEP=UserEP+"&Forum_user(7)&",UserCP=UserCP+"&Forum_user(12)&" where userid="&rs(2))
			set ars=conn.execute("select count(*) from "&NowUseBBS&" where locktopic=3 and rootid="&rs(0))
			haveaudit=ars(0)
			if isnull(haveaudit) or haveaudit=0 then conn.execute("update "&NowUseBBS&" set isaudit=0 where rootid="&rs(0))
		else
			set trs=conn.execute("select rootid,dateandtime,PostUserID from "&NowUseBBS&" where Announceid="&id)
			if not (trs.eof and trs.bof) then
			'更新主题最后回复数据和回复数
			bbsnum=bbsnum+1
			topicnum=topicnum+1
			if datediff("d",trs(1),Now())=0 then todaynum=todaynum+1
			conn.execute("update "&NowUseBBS&" set locktopic=0,isaudit=0 where Announceid="&id)
			conn.execute("update [user] set article=article+1,userWealth=userWealth+"&Forum_user(2)&",UserEP=UserEP+"&Forum_user(7)&",UserCP=UserCP+"&Forum_user(12)&" where userid="&trs(2))
			set ars=conn.execute("select count(*) from "&NowUseBBS&" where locktopic=3 and rootid="&trs(0))
			haveaudit=ars(0)
			if isnull(haveaudit) or haveaudit=0 then conn.execute("update "&NowUseBBS&" set isaudit=0 where rootid="&trs(0))
			IsEndReply(trs(0))
			end if
		end if
	end if
next
set rs=nothing
'更新论坛总数据和版面数据
if cint(request("actiontype"))=1 then
	update boardid,bbsnum,topicnum,todaynum
end if
sucmsg="<br><li>操作成功！"
call dvbbs_suc()
end sub

function IsEndReply(TopicID)
isEndReply=false
dim trs
dim LastPostInfo,iTotalUseTable
dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
dim LastPost,uploadpic_n,LastPostUserID,LastID
set trs=conn.execute("select LastPost,PostTable from Topic where Topicid="&Topicid)
if not (trs.eof and trs.bof) then
	LastPostInfo=split(trs(0),"$")
	iTotalUseTable=trs(1)
end if
set trs=conn.execute("select top 1 topic,body,Announceid,dateandtime,username,PostUserid,rootid from "&iTotalUseTable&" where rootid="&TopicID&" and locktopic<2 order by Announceid desc")
if not(trs.eof and trs.bof) then
	body=trs(1)
	LastRootid=trs(2)
	LastPostTime=trs(3)
	LastPostUser=replace(trs(4),"$","")
	LastTopic=left(replace(body,"$",""),20)
	LastPostUserID=trs(5)
	LastID=trs(6)
else
	LastTopic="无"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="无"
	LastPostUserID=0
	LastID=0
end if
LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & replace(left(replace(LastTopic,"'",""),20),"$","") & "$" & LastPostInfo(4) & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
conn.execute("update topic set LastPost='"&LastPost&"',child=child+1,LastPostTime='"&LastPostTime&"' where topicid="&TopicID)
set trs=nothing
end function


'更新论坛总数据和版面数据
function update(boardid,bbsnum,topicnum,todaynum)
dim lastpost_1,trs
dim LastTopic,LastRootid,LastPostTime,LastPostUser
dim LastPost,uploadpic_n,Lastpostuserid,Lastid
dim UpdateBoardID
'本论坛和上级论坛ID
UpdateBoardID=BoardParentStr & "," & BoardID
'版面最后回复数据
set trs=conn.execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&NowUseBBS&" b inner join Topic T on b.rootid=T.TopicID where b.boardid="&boardid&" and b.locktopic<2 order by b.announceid desc")
if not(trs.eof and trs.bof) then
	Lasttopic=replace(left(replace(trs(0),"'",""),15),"$","")
	LastRootid=trs(1)
	LastPostTime=trs(2)
	LastPostUser=trs(3)
	LastPostUserid=trs(4)
	Lastid=trs(5)
else
	LastTopic="无"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="无"
	LastPostUserid=0
	Lastid=0
end if
set trs=nothing
LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
'总版面最后回复数据
set trs=conn.execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&NowUseBBS&" b inner join Topic T on b.rootid=T.TopicID where b.locktopic<2 order by b.announceid desc")
if not(trs.eof and trs.bof) then
	Lasttopic=replace(left(replace(trs(0),"'",""),15),"$","")
	LastRootid=trs(1)
	LastPostTime=trs(2)
	LastPostUser=trs(3)
	LastPostUserid=trs(4)
	Lastid=trs(5)
else
	LastTopic="无"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="无"
	LastPostUserid=0
	Lastid=0
end if
LastPost_1=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID

Dim SplitUpBoardID,SplitLastPost
SplitUpBoardID=split(UpdateBoardID,",")
For i=0 to ubound(SplitUpBoardID)
	set trs=conn.execute("select LastPost from board where boardid="&SplitUpBoardID(i))
	if not (trs.eof and trs.bof) then
	SplitLastPost=split(trs(0),"$")
	if isnull(SplitLastPost(1)) then SplitLastPost(1)=0
	if ubound(SplitLastPost)=7 and clng(LastRootID)<>clng(SplitLastPost(1)) then
		conn.execute("update board set LastPost='"&LastPost&"' where boardid="&SplitUpBoardID(i))
	end if
	end if
Next
conn.execute("update board set  LastBbsNum=Lastbbsnum+"&bbsnum&",LastTopicNum=LastTopicNum+"&TopicNum&",TodayNum=TodayNum+"&todaynum&" where boardid in ("&UpdateBoardID&")")
conn.execute("update config set  BbsNum=bbsnum+"&bbsnum&",TopicNum=TopicNum+"&TopicNum&",TodayNum=TodayNum+"&todaynum&",LastPost='"&LastPost_1&"' where active=1")
set trs=nothing
end function
%>