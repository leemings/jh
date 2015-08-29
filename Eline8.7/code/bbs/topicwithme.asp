<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	dim totalrec
	dim n
	dim orders,ordername
	dim currentpage,page_count,Pcount
	currentPage=request("page")
	if not founduser then
		Errmsg="您还没有登录，请<a href=login.asp>登录</a>后浏览"
		founderr=true
	end if
	orders=request("s")
	if orders="" or not isInteger(orders) then
		orders=1
	else
		orders=clng(orders)
	end if
	if orders=1 then
		ordername="我参与的主题"
		sql="select top 200 * from topic where topicid in (select top 200 rootid from "&NowUseBBS&" where postuserid="&userid&" and rootid<>0 order by announceid desc) and locktopic<>2 order by topicid desc"
	elseif orders=2 then
		sql="select top 200 * from topic where postuserid="&userid&" and child>0 and locktopic<>2 ORDER BY topicid desc"
		ordername="已被回复的我的发言"
	else
		ordername="我发表的主题"
		sql="select top 200 * from topic where postuserid="&userid&" and locktopic<>2 order by topicid desc"
	end if	
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if
	stats=ordername
if founderr then
	call nav()
	call head_var(2,0,"","")
else
	call nav()
	call head_var(2,0,"","")
	call AnnounceList1()
	call listPages3()
end if
call activeonline()
call footer()

sub AnnounceList1()
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		call showEmptyBoard1()
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		call showpagelist1()
	end if
	end sub

	REM 显示帖子列表	
	sub showPageList1()
	i=0
	response.write "<p align=center>当前论坛数据很可能进行了分表保存数据<BR>当前仅显示您在当前数据表中最前200帖，如果您需要查询更多数据请到搜索页面</p>"%>
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
	<%response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
	response.write "<TR align=middle>"
	response.write "<Th height=27 width=32 id=tabletitlelink>状态</th>"
	response.write "<Th width=*>主 题  (点<img src="&Forum_info(7)&"plus.gif>即可展开帖子列表)</Th>"
	response.write "<Th width=80>作 者</Th>"
	response.write "<Th width=64>回复/人气</Th>"
	response.write "<Th width=70>最后更新</Th>"
	response.write "<Th width=80>回复人</Th>"
	response.write "</TR>"
	dim Ers, Eusername, Edateandtime,body
	dim vrs,votenum,pnum,p,iu,votenum_1
	while (not rs.eof) and (not page_count = rs.PageSize)
		boardid=rs("boardid")
		response.write "<TR align=middle>"
		response.write "<TD class=tablebody2 width=32 height=27>"
		response.write TopicIcon(rs("BoardID"),rs("topicid"),rs("lastposttime"),rs("istop"),rs("isVote"),rs("LockTopic"),iif(Rs("child")>=Cint(forum_setting(44)),1,0))
		response.write "</TD>"
		response.write "<TD onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1' align=left  class=tablebody1 width=*>"
		if Rs("child")=0 then
			response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& Rs("topicid") &""">"
		else
			response.write "<img loaded=""no"" src="""& Forum_info(7)&Forum_boardpic(15)&""" id=""followImg"& Rs("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& Rs("topicid") &","& Rs("boardid") &")"" title=展开帖子列表>"
		end if
		'response.write "&nbsp;<img src='"& Forum_info(8) & rs("Expression") &"'>&nbsp;"
		response.write "<a href=""dispbbs.asp?boardID="& rs("boardid") &"&ID="& rs("topicid") &""">"
		if len(rs("title"))>26 then
			response.write ""&left(htmlencode(rs("title")),26)&"..."
		else
			response.write htmlencode(rs("title"))
		end if
		response.write "</a>"
		if (rs("child")+1)>cint(Forum_Setting(11)) then
			response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&Forum_statePic(6)&"><b>"
			Pnum=(Cint(rs("child")+1)/cint(Forum_Setting(11)))+1
			for p=1 to Pnum
			response.write " <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&P&"'><FONT color="&Forum_body(8)&"><b>"&p&"</b></font></a> "
			next
			response.write "</b>]"
		end if
		if rs("isbest")=1 then
			response.write "&nbsp;&nbsp;<img src="""&Forum_info(7)&Forum_statePic(4)&""">"
		end if
		response.write "</TD>"
		response.write "<TD class=tablebody2 width=80><a href=dispuser.asp?name="& rs("postusername") &">"& rs("postusername") &"</a></TD>"
		response.write "<TD class=tablebody1 width=64>"
		response.write ""& rs("child") &"/"& rs("hits") &"</TD>"
		response.write "<TD align=center class=tablebody2 width=70>"&month(rs("dateandtime"))&"-"&day(rs("dateandtime"))&"&nbsp;"&FormatDateTime(rs("dateandtime"),4)&"</TD>"
		response.write "<TD align=center class=tablebody1 width=80>"
		if rs("child")> 0 then
			response.write split(rs("lastpost"),"$")(0)
		else
			response.write "------"
		end if
		response.write "</TD></TR>"
		response.write "<tr style=""display:none"" id=""follow"& Rs("topicid") &"""><td colspan=6 id=""followTd"& Rs("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& Rs("topicid") &","&Rs("boardid")&")"">正在读取关于本主题的跟帖，请稍侯……</div></td></tr>"
	  page_count = page_count + 1
	rs.movenext
	wend
	response.write "</table>"
	end sub

	sub listPages3()
	dim endpage
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<strong>"&Forum_Setting(11)&"</b> 帖数<b>"&totalrec&"</b></td>"&_
			"<td valign=middle nowrap><div align=right><p>分页： "
	call DispPageNum(currentpage,PCount,"""?page=","&s="&orders&"""")
	response.write "</p></div></font></td></tr></table>"
	rs.close
	set rs=nothing	
	end sub 

	sub showEmptyBoard1()
%>
<TABLE cellPadding=4 cellSpacing=1 class=tableborder1 align=center>
<TBODY>
<TR align=middle>
<Th height=25>状态</Th>
<Th>主 题  (点心情符为开新窗浏览)</Th>
<Th>作 者</Th> 
<Th>回复/人气</Th> 
<Th>最新回复</Th>
</TR> 
<tr><td colSpan=5 width=100% class=tablebody1>暂无内容，欢迎发帖：）</td></tr>
</TBODY></TABLE>
<%
	end sub
%>