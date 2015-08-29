<marquee scrollamount=3 onmouseover=this.stop(); onmouseout=this.start();><%
	sql="select top 5 title,addtime from bbsnews where boardid="&boardid&"  order by id desc"
	set rs=conn.execute(sql)
	if rs.bof and rs.eof then
		response.write "<a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='当前没有公告'><b>当前没有公告</b></ACRONYM><a> ("& now() &")"
	else
		do while (not rs.eof)
			response.write "<a href='announcements.asp?boardid="& boardid &"' target=_blank><b>"& rs("title") &"</b></a> ("& rs("addtime") &")    "
			rs.movenext 
		loop
		response.write "</font>"
	end if
	set rs=nothing
%></marquee>
