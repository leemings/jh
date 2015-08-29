<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp"-->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	call main()
end if
conn.close
set conn=nothing

sub main()%>
	<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
		<tr>
			<td><b>浏览上传文件</b>：本功能必须服务器支持FSO权限方能使用，FSO使用帮助请浏览微软网站。文件目录为<%=downpath%>，如果您服务器不支持FSO请手动管理。</td>
		</tr>
		<tr> 
			<td><font color=<%=forum_body(8)%>>注意：如果您删除此文件，那么数据库中关于此文件的软件相关信息也会被删除，请小心使用！</font></td>
		</tr>
		<tr> 
			<td><b>当前浏览 <%=downpath%> 文件目录的所有文件</b>：</td>
		</tr>
		<tr> 
			<td><%
				dim objFSO
				dim uploadfolder
				dim uploadfiles
				dim upname
				dim upfilename
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				if request("filename")<>"" then
					objFSO.DeleteFile(Server.MapPath(downpath&request("filename")))
				end if
				set rss=conndown.execute("select * from download where filename ='"&request("filename")&"' or filename1 = '"&request("filename")&"' or filename2= '"&request("filename")&"'")
				if not (rss.bof and rss.eof) then 
					sql="delete from download where filename = '"&request("filename")&"' or filename1 = '"&request("filename")&"' or filename2 = '"&request("filename")&"'"
					conndown.execute sql
				end if
				Set uploadFolder=objFSO.GetFolder(Server.MapPath(downpath))
				Set uploadFiles=uploadFolder.Files
				For Each Upname In uploadFiles
					upfilename=downpath&upname.name
					response.write "<a href="""&upfilename&""">"&upfilename&"</a> <a href='?path="&downpath&"&filename="&upname.name&"'>删除</a><br>"
				next
				set uploadFolder=nothing
				set uploadFiles=nothing
			%></td>
		</tr>
	</table>
<%end sub%>







