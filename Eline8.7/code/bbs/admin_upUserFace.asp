<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
dim admin_flag

dim objFSO
dim uploadfolder
dim uploadfiles
dim upname
dim uid,faceid
dim usernames
dim userface,dnum

dim upfilename
dim pagesize, page,filenum, pagenum


admin_flag="71"
if not master or instr(session("flag"),admin_flag)=0 then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_Error()
else
	call main()
end if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr>
<td valign=top>
注意：本功能需要主机开放FSO权限，FSO相关帮助请看微软帮助文档<BR>
在这里您可以管理论坛所有用户自定义头像上传文件，搜索用户头像请用用户ID进行搜索<BR>
用户ID的获得可以通过用户信息管理中搜索相关用户，然后将鼠标移到用户名连接上，查看连接属性，参数UserID=后面既是用户的ID
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder" style="table-layout:fixed;word-break:break-all">
<tr align=center><th width="*" height=25>文件名</th><th width="100">所属用户</th><th width="50">大小</th><th width="120">最后访问</th><th width="120">上传日期</th><th width="35">管理</th></tr>
<form method="POST" action="?action=delall">
<%

pagesize=20
page=request.querystring("page")
if page="" or not isnumeric(page) then
	page=1
else
	page=int(page)
end if

if trim(request("action"))<>"" then
	if trim(request("action"))="delall" then
	call delface()
	else
	call maininfo()
	end if
else
	call maininfo()
end if

call foot()
end sub

sub maininfo()


Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("filename")<>"" then
objFSO.DeleteFile(Server.MapPath("uploadFace\"&request("filename")))
end if
Set uploadFolder=objFSO.GetFolder(Server.MapPath("uploadFace\"))
Set uploadFiles=uploadFolder.Files

filenum=uploadfiles.count
pagenum=int(filenum/pagesize)
if filenum mod pagesize>0 then
	pagenum=pagenum+1
end if
if page> pagenum then
	page=1
end if
i=0
For Each Upname In uploadFiles
	i=i+1
	if i>(page-1)*pagesize and i<=page*pagesize then
	upfilename="uploadFace/"&upname.name

	if instr(upname.name,"_") then    '取出头像的用户名
		uid=split(upname.name,"_")
		faceid=uid(0)
		if  isInteger(faceid)	then
		set rs=conn.execute("select username from [user] where   userid="&faceid&"  ")
		if not rs.eof  then
		usernames=rs(0)
		end if
		rs.close
		set rs=nothing
		end if		
	end if

		response.write "<tr><td class=forumRow height=23><a href=""uploadface/"&upname.name&""" target=_blank>"&upname.name&"</a></td>"
		response.write "<td align=right class=forumRowHighlight>"&usernames&"</td>"
		response.write "<td align=right class=forumRow>"& upname.size &"</td>"
		response.write "<td align=center class=forumRowHighlight>"& upname.datelastaccessed &"</td>"
		response.write "<td align=center class=forumRow>"& upname.datecreated &"</td>"
		response.write "<td align=center class=forumRowHighlight><a href='?filename="&upname.name&"'>删除</a></td></tr>"

	elseif i>page*pagesize then
		exit for
	end if
	usernames=""
next

end sub 

sub delface()
dnum=0
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("filename")<>"" then
objFSO.DeleteFile(Server.MapPath("uploadFace\"&request("filename")))
end if
Set uploadFolder=objFSO.GetFolder(Server.MapPath("uploadFace\"))
Set uploadFiles=uploadFolder.Files

filenum=uploadfiles.count
pagenum=int(filenum/pagesize)
if filenum mod pagesize>0 then
	pagenum=pagenum+1
end if
if page> pagenum then
	page=1
end if
i=0
For Each Upname In uploadFiles
	i=i+1
	if i>(page-1)*pagesize and i<=page*pagesize then
	upfilename="uploadFace/"&upname.name

if instr(upname.name,"_") then    '取出头像的用户名
		uid=split(upname.name,"_")
		faceid=uid(0)
		if  isInteger(faceid)	then
			set rs=conn.execute("select username,face from [user] where   userid="&faceid&"  ")
			if not rs.eof  then
			usernames=rs(0)
			userface=trim(rs(1))
				if instr(upfilename,userface)=0 then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write "头像已更改,文件："&upfilename&"已删除<br>"&userface
			dnum=dnum+1
				end if
		else
		objFSO.DeleteFile(Server.MapPath(upfilename))
		response.write "用户已注销,文件："&upfilename&"已删除<br>"
		dnum=dnum+1
		end if
		rs.close
		set rs=nothing
		end if
else
'清理没有用户ID的头像文件
			sql="select top 1 userid from [user] where face like '%"&upfilename&"%' "
			set rs=conn.execute(sql)
			if rs.eof then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write "已清查删除文件："&upfilename&"<br>"
			dnum=dnum+1
			end if
			rs.close
			set rs=nothing
end if
	elseif i>page*pagesize then
		exit for
	end if
next
response.write " 共清理 "&dnum&" 个文件  "
end sub

sub foot()
set uploadFolder=nothing
set uploadFiles=nothing
%>
<tr><td colspan=6 class=forumRow height=30>
<%
if page>1 then
	response.write "<a href=?page=1>首页</a>&nbsp;&nbsp;<a href=?page="& page-1 &">上一页</a>&nbsp;&nbsp;"
else
	response.write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
end if
if page<i/pagesize then
	response.write "<a href=?page="& page+1 &">下一页</a>&nbsp;&nbsp;<a href=?page="& pagenum &">尾页</a>"
else
	response.write "下一页&nbsp;&nbsp;尾页"
end if
%>
<input type="submit" value="清理">
</td><tr></form>
</table>
<br>
<% end sub%>
