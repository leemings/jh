<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
dim action
dim admin_flag
action=trim(request("action"))

dim dbpath,bkfolder,bkdbname,fso,fso1

select case action
case "CompressData"		'压缩数据
	admin_flag="61"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		call CompressData()

	end if	

case "BackupData"		'备份数据
	admin_flag="62"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("act")="Backup" then
		call updata()
		else
		call BackupData()
		end if
	end if

case "RestoreData"		'恢复数据
	admin_flag="63"
	dim backpath
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
			response.write "请输入您要恢复成的数据库全名"	
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
		
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			response.write "成功恢复数据！"
			else
			response.write "备份目录下并无您的备份文件！"	
			end if
		else
		call RestoreData()
		end if
	end if

 case "SpaceSize"		'系统空间占用
	admin_flag="64"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else	
		call SpaceSize()
	end if

case else
		Errmsg=Errmsg+"<br>"+"<li>选取相应的操作。"
		call dvbbs_error()

end select

conn.close
set conn=nothing
response.write"</body></html>"


'====================系统空间占用=======================
sub SpaceSize()
on error resume next
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;系统空间占用情况
  					</th>
  				</tr> 	
 				<tr>
 					<td class="forumrow"> 			
 			<blockquote>
 			<br> 			
 			法规数据占用空间：&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("data")%> height=10>&nbsp;<%showSpaceinfo("data")%><br><br>
 			备份数据占用空间：&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("databackup")%> height=10>&nbsp;<%showSpaceinfo("databackup")%><br><br>
 			程序文件占用空间：&nbsp;<img src="pic/bar1.gif" width=<%=drawspecialbar%> height=10>&nbsp;<%showSpecialSpaceinfo("Program")%><br><br>
 			心情图片占用空间：&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("images")%> height=10>&nbsp;<%showSpaceinfo("face")%><br><br>
 			系统图片占用空间：&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("pic")%> height=10>&nbsp;<%showSpaceinfo("pic")%><br><br>
 			上传头像占用空间：&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("uploadFace")%> height=10>&nbsp;<%showSpaceinfo("uploadFace")%><br><br>
 			上传文件占用空间：&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("uploadFile")%> height=10>&nbsp;<%showSpaceinfo("uploadFile")%><br><br>	
 			系统占用空间总计：<br><img src="pic/bar1.gif" width=400 height=10> <%showspecialspaceinfo("All")%>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
end sub

'====================恢复数据库=========================
sub RestoreData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder"		<tr>
	<th height=25 >
   					&nbsp;&nbsp;<B>恢复论坛数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
				<form method="post" action="ADMIN_data.asp?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;备份数据库路径(相对)：<input type=text size=30 name=DBpath value="DataBackup\eline_bbs_6.3.0.asp">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;目标数据库路径(相对)：<input type=text size=30 name=backpath value="<%=db%>"><BR>&nbsp;&nbsp;填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确），然后修改conn.asp文件，如果目标文件名和当前使用数据库名一致的话，不需修改conn.asp文件<BR>
						&nbsp;&nbsp;<input type=submit value="恢复数据"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认备份数据库文件为DataBackup\DVBBS60.MDB，请按照您的备份文件自行修改。<br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

'====================备份数据库=========================
sub BackupData()
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>备份论坛数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_data.asp?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;
						当前数据库路径(相对路径)：<input type=text size=15 name=DBpath value="<%=db%>"><BR>&nbsp;&nbsp;
						备份数据库目录(相对路径)：<input type=text size=15 name=bkfolder value=Databackup>&nbsp;如目录不存在，程序将自动创建<BR>&nbsp;&nbsp;
						备份数据库名称(填写名称)：<input type=text size=15 name=bkDBname value=DVBBS60.MDB>&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
						&nbsp;&nbsp;<input type=submit value="确定"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认数据库文件为Data\DVBBS60.MDB，<B>请一定不能用默认名称命名备份数据库</B><br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径				</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

sub updata()
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			response.write "备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname
		Else
			response.write "找不到您所需要备份的文件。"
		End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================压缩数据库 =========================
sub CompressData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<form action="Admin_data.asp?action=CompressData" method="post">
<tr>
<td class="forumrow" height=25><b>注意：</b><br>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作） </td>
</tr>
<tr>
<td class="forumrow">压缩数据库：<input type="text" name="dbpath" value=Data\DVBBS60.MDB>&nbsp;
<input type="submit" value="开始压缩"></td>
</tr>
<tr>
<td class="forumrow"><input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br><br></td>
</tr>
<form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "你的数据库, " & dbpath & ", 已经压缩成功!" & vbCrLf

Else
	CompactDB = "数据库名称或路径不正确. 请重试!" & vbCrLf
End If

End Function


'=====================系统空间参数=========================
	Sub ShowSpaceInfo(drvpath)
 		dim fso,d,size,showsize
 		set fso=server.createobject("scripting.filesystemobject") 		
 		drvpath=server.mappath(drvpath) 		 		
 		set d=fso.getfolder(drvpath) 		
 		size=d.size
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	End Sub	
 	
 	Sub Showspecialspaceinfo(method)
 		dim fso,d,fc,f1,size,showsize,drvpath 		
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpath=server.mappath("pic")
 		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
 		set d=fso.getfolder(drvpath) 		
 		
 		if method="All" then 		
 			size=d.size
 		elseif method="Program" then
 			set fc=d.Files
 			for each f1 in fc
 				size=size+f1.size
 			next	
 		end if	
 		
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	end sub 	 	 	
 	
 	Function Drawbar(drvpath)
 		dim fso,drvpathroot,d,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		drvpath=server.mappath(drvpath) 		
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		
 		barsize=cint((size/totalsize)*400)
 		Drawbar=barsize
 	End Function 	
 	
 	Function Drawspecialbar()
 		dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		set fc=d.files
 		for each f1 in fc
 			size=size+f1.size
 		next	
 		
 		barsize=cint((size/totalsize)*400)
 		Drawspecialbar=barsize
 	End Function 	
%>