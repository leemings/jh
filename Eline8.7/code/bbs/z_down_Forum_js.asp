<script Language="JavaScript">
<%dim rsclass,subrs
set rs= server.createobject ("adodb.recordset")
sql="select class,classid from Dclass"
rs.open sql,conndown,1,1
if not (rs.bof and rs.eof) then
	do while not rs.eof
		response.write "var classlist"&rs("classid")&"= '"
		set rsclass= server.createobject ("adodb.recordset")
		sql="select Nclass,Nclassid,classid from DNclass where classid="&rs("classid")
		rsclass.open sql,conndown,1,1
		if rsclass.eof and rsclass.bof then
			response.write "无子分类"
		else
			do while not rsclass.eof
				response.write "<a style=font-size:9pt;line-height:12pt; href=\'z_down_index.asp?classid="&rs("classid")&"&Nclassid="&rsclass("Nclassid")&"\'>"&rsclass("Nclass")&"</a><br>"
				rsclass.movenext
			loop
		end if
		rsclass.close
		response.write "';"&vbNewLine
		rs.movenext
	loop
end if
rs.close
rs.open "select * from [Dclass] order by classID asc",conndown,1,1
if not(rs.bof and rs.eof) then
	do while not rs.eof 
		Response.Write "var catemenu"&rs("classID")&"='"
		set SubRs= server.createobject ("adodb.recordset")
		SubRs.open "select * from [DNclass] where ClassID="&Rs("ClassID")&" order by NclassID asc",conndown,1,1
		if SubRs.bof and SubRs.eof then
			response.write "没有添加小类"
		else
			do while not SubRs.eof 		
				response.write "<a style=font-size:9pt;line-height:12pt; href=\""z_admin_down_adminedit.asp?classID="&SubRs("classID")&"&NclassID="&SubRs("NclassID")&"\"">"&SubRs("NClass")&"</a><br>"
				SubRs.movenext
			loop
		end if
		SubRs.close
		Response.Write "';" & vbcrlf
		rs.movenext
	loop
end if
rs.close
rs.open "select * from [Dclass] order by classID asc",conndown,1,1
if not(rs.bof and rs.eof) then
	do while not rs.eof 
		Response.Write "var catemenunew"&rs("classID")&"='"
		set SubRs= server.createobject ("adodb.recordset")
		SubRs.open "select * from [DNclass] where ClassID="&Rs("ClassID")&" order by NclassID asc",conndown,1,1
		if SubRs.bof and SubRs.eof then
			response.write "没有添加小类"
		else
			do while not SubRs.eof 		
				response.write "<a style=font-size:9pt;line-height:12pt; href=\""z_down_adminedit.asp?classID="&SubRs("classID")&"&NclassID="&SubRs("NclassID")&"\"">"&SubRs("NClass")&"</a><br>"
				SubRs.movenext
			loop
		end if
		SubRs.close
		Response.Write "';" & vbcrlf
		rs.movenext
	loop
end if
rs.close%>
</script>