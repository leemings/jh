<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
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
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_Error()
else
	call main()
end if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr>
<td valign=top>
ע�⣺��������Ҫ��������FSOȨ�ޣ�FSO��ذ����뿴΢������ĵ�<BR>
�����������Թ�����̳�����û��Զ���ͷ���ϴ��ļ��������û�ͷ�������û�ID��������<BR>
�û�ID�Ļ�ÿ���ͨ���û���Ϣ��������������û���Ȼ������Ƶ��û��������ϣ��鿴�������ԣ�����UserID=��������û���ID
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder" style="table-layout:fixed;word-break:break-all">
<tr align=center><th width="*" height=25>�ļ���</th><th width="100">�����û�</th><th width="50">��С</th><th width="120">������</th><th width="120">�ϴ�����</th><th width="35">����</th></tr>
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

	if instr(upname.name,"_") then    'ȡ��ͷ����û���
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
		response.write "<td align=center class=forumRowHighlight><a href='?filename="&upname.name&"'>ɾ��</a></td></tr>"

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

if instr(upname.name,"_") then    'ȡ��ͷ����û���
		uid=split(upname.name,"_")
		faceid=uid(0)
		if  isInteger(faceid)	then
			set rs=conn.execute("select username,face from [user] where   userid="&faceid&"  ")
			if not rs.eof  then
			usernames=rs(0)
			userface=trim(rs(1))
				if instr(upfilename,userface)=0 then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write "ͷ���Ѹ���,�ļ���"&upfilename&"��ɾ��<br>"&userface
			dnum=dnum+1
				end if
		else
		objFSO.DeleteFile(Server.MapPath(upfilename))
		response.write "�û���ע��,�ļ���"&upfilename&"��ɾ��<br>"
		dnum=dnum+1
		end if
		rs.close
		set rs=nothing
		end if
else
'����û���û�ID��ͷ���ļ�
			sql="select top 1 userid from [user] where face like '%"&upfilename&"%' "
			set rs=conn.execute(sql)
			if rs.eof then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write "�����ɾ���ļ���"&upfilename&"<br>"
			dnum=dnum+1
			end if
			rs.close
			set rs=nothing
end if
	elseif i>page*pagesize then
		exit for
	end if
next
response.write " ������ "&dnum&" ���ļ�  "
end sub

sub foot()
set uploadFolder=nothing
set uploadFiles=nothing
%>
<tr><td colspan=6 class=forumRow height=30>
<%
if page>1 then
	response.write "<a href=?page=1>��ҳ</a>&nbsp;&nbsp;<a href=?page="& page-1 &">��һҳ</a>&nbsp;&nbsp;"
else
	response.write "��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
end if
if page<i/pagesize then
	response.write "<a href=?page="& page+1 &">��һҳ</a>&nbsp;&nbsp;<a href=?page="& pagenum &">βҳ</a>"
else
	response.write "��һҳ&nbsp;&nbsp;βҳ"
end if
%>
<input type="submit" value="����">
</td><tr></form>
</table>
<br>
<% end sub%>
