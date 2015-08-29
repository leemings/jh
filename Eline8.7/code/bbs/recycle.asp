<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
dim bBoardEmpty
dim totalrec
dim n,RowCount
dim p
dim currentpage,page_count,Pcount
dim tablename
if not membername<>"" then
	Errmsg=Errmsg+"<br><li>您没有权限浏览本页。"
	founderr=true
end if

stats="论坛回收站"
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

call nav()
call head_var(2,0,"","")
if founderr then
	call dvbbs_error()
else
	call AnnounceList1()
	call listPages3()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub AnnounceList1()
	response.write "<table cellpadding=0 cellspacing=0 border=0 width="&Forum_body(12)&" align=center><tr><td align=center width=2 valign=middle>&nbsp;</td><td align=left valign=middle> 本页面只有系统管理员可进行操作，请选择需要的操作，列出所有：<a href=?tablename=topic>主题表数据</a>"
	for i=0 to ubound(allposttable)
		response.write " | <a href=?tablename="&allposttable(i)&">帖子"&allposttablename(i)&"</a>"
	next
	if master then
		response.write "<BR>注意：还原主题表数据将连跟帖（帖子）表数据一起还原，删除主题表数据将连跟帖（帖子）表一起删除</td><td align=right>&nbsp;</td></tr></table><BR>"
	else
		response.write "</td><td align=right>&nbsp;</td></tr></table><BR>"
	end if

	if instr(request("tablename"),"bbs")>0 then
		tablename=replace(request("tablename"),"'","")
		tablename=replace(tablename,";","")
		tablename=replace(tablename,"--","")
		sql="select AnnounceID,boardID,UserName,Topic,body,DateAndTime from "&tablename&" where locktopic=2 and not parentid=0 order by announceid desc"
	else
		sql="select topicID,boardID,PostUserName,Title,title as body,DateAndTime from topic where locktopic=2 order by topicid desc"
		tablename="topic"
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		'论坛无内容
		call showEmptyBoard1()
	else
		rs.PageSize = cint(Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		call showpagelist1() 
	end if
end sub

REM 显示帖子列表	
sub showPageList1()
	dim body
	dim vrs
	dim votenum,votenum_1
	dim pnum
	i=0
	response.write "<form name=recycle action=admin_recycle.asp method=post><input type=hidden value="&tablename&" name=tablename>"
	response.write "<TABLE cellPadding=1 cellSpacing=1 align=center class=tableborder1><TBODY><TR align=middle height=25>"
	if master then
		response.write "<Th width=32>状态</Th>"
	end if
	response.write "<Th width=*>主 题</Th><Th width=80>作 者</Th><Th width=70>最后更新</th><th width=80>回复人</Th></TR>"
	while (not rs.eof) and (not page_count = rs.PageSize)
		response.write "<TR align=middle height=27>"
		if master then
			response.write "<TD class=tablebody2 width=32><input type=checkbox name=topicid value="&rs(0)&"></TD>"
		end if
		response.write "<TD onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1' align=left class=tablebody1 width=*>&nbsp;"
		body=left(htmlencode(replace(reUBBCode(rs(4),false),chr(10)," ")),26)&"..."
		response.write "<a href=""javascript:openScript('viewrecycle.asp?id="&rs(0)&"&tablename="&tablename&"',500,400)"">"
		if rs(3)="" or isnull(rs(3)) then
			response.write body
		else
			response.write htmlencode(reUBBCode(rs(3),false))
		end if
		response.write "</a></TD>"
		response.write "<TD class=tablebody2 width=80><a href=dispuser.asp?name="& rs(2) &">"& rs(2) &"</a></TD>"
		response.write "<TD class=tablebody1 width=70>"&month(rs(5))&"-"&day(rs(5))&"&nbsp;"&FormatDateTime(rs(5),4)&"</td>"
		response.write "<TD class=tablebody1 width=80>------</TD>"
		response.write "</TR>"
	  page_count = page_count + 1
		rs.movenext
	wend
end sub

sub listPages3()
	dim endpage
	Pcount=rs.PageCount
	if master and totalrec>0 then
		response.write "<TR height=27><TD class=tablebody2 colspan=5><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">选中所有显示帖子&nbsp;<input type=submit name=action onclick=""{if(confirm('确定还原选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=还原>"
		'if session("flag")<>"" then
			response.write "&nbsp;<input type=submit name=action onclick=""{if(confirm('确定删除选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=删除>&nbsp;<input type=submit name=action onclick=""{if(confirm('确定清除回收站所有的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=清空回收站>"
		'end if
		response.write "</TD></TR>"
	end if
	response.write "</TBODY></TABLE></form>"
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center><tr><td valign=middle nowrap>页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页	每页<b>"&Forum_Setting(11)&"</b> 帖数<b>"&totalrec&"</b></td><td valign=middle nowrap><div align=right><p>分页： "
	call DispPageNum(currentpage,PCount,"""?page=","&tablename="&tablename&"""")
	response.write "</p></div></td></tr></table>"
	rs.close
	set rs=nothing
end sub 

sub showEmptyBoard1()
	response.write "<TABLE class=tableborder1 cellPadding=4 cellSpacing=1 align=center><TBODY><TR align=middle><Th height=25>状态</Th><Th>主 题</Th><Th>作 者</Th> <Th>回复/人气</Th> <Th>最新回复</Th></TR> <tr><td colSpan=5 width=100% class=tablebody1>论坛回收站暂无内容。</td></tr></TBODY></TABLE>"
end sub
%>
<script language="JavaScript">
<!--
function CheckAll(form)
{
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name != 'chkall')
			e.checked = form.chkall.checked; 
	}
}
//-->
</script>
