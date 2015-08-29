<!--#include FILE="conn.asp"-->
<!--#include FILE="upload.inc"-->
<!--#include file="inc/const.asp" -->
<html>
<head>
<title>文件上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body topmargin="0" leftmargin="0">
<script>
parent.document.all.frmAnnounce.Submit.disabled=false;
parent.document.all.frmAnnounce.Submit2.disabled=false;
</script>
<table width="100%" border=0 cellspacing=0 cellpadding=0>
	<tr>
		<td class=tablebody2 height=40 width=100% valign=middle><%
			Server.ScriptTimeOut=999999'要是你的论坛支持上传的文件比较大，就必须设置。
			'上传方式upload_type值: 0＝无组件，1＝lyfupload
			dim upload_type
			upload_type=0
			'创建生成预览图片,需要CreatePreviewImage组件支持,upload_view值: 0=不支持,1=支持(根目录下要有PreviewImage文件夹存放文件)
			dim upload_view
			upload_view=0
			'定义变量
			dim Forumupload,ranNum
			dim formName,formPath,filename,file_name,fileExt,Filesize,F_Type
			dim upNum,dateupnum
			dim rename,DownloadID
			dim previewpath,F_Viewname
			F_Viewname=""
			previewpath="PreviewImage/"
			upNum=request.cookies("upNum")
			if upnum ="" then upnum=0
			upNum=int(upNum)
			dateupnum=request.cookies("dateupnum")
			if dateupnum ="" then dateupnum=0
			dateupnum=int(dateupnum)
			if Cint(GroupSetting(7))=0 then
				response.write "您没有在本论坛上传文件的权限"
			elseif not founduser then
				response.write "只有登录用户方能上传！"
			elseif upNum >= int(GroupSetting(40)) then
				response.write "一次只能上传"&GroupSetting(40)&"个文件！"
			elseif dateupnum >= int(GroupSetting(50)) then
				response.write "您今天上传的文件已超出了"&GroupSetting(50)&"个！"
			else
				select case upload_type
				case 0
					call upload_0()
				case 1
					call upload_1()
				case else
					response.write "本系统未开放上传功能"
				end select
			end if
		%></td>
	</tr>
</table>
</body>
</html>
<%'===========================无组件上传============================
sub upload_0()
	dim upload,file
	set upload=new Upload_file ''建立上传对象
	upload.GetData (GroupSetting(44)*1024)   '取得上传数据,不限大小
	if upload.err > 0 then
		select case upload.err
		case 1
			Response.Write "请先选择你要上传的文件　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
		case 2
			Response.Write "文件大小超过了限制 "&GroupSetting(44)&"K　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
		end select
		exit sub
	else
		dim curamount,totalamount
		formPath=upload.form("filepath")
		'在目录后加(/)
		if right(formPath,1)<>"/" then formPath=formPath&"/"
		for each formName in upload.file ''列出所有上传了的文件
			if upNum >= int(GroupSetting(40)) or dateupnum >= clng(GroupSetting(50)) then
				response.write "已达到上传数的上限。"
				exit sub
			end if
			set file=upload.file(formName)  ''生成一个文件对象
			fileExt=lcase(file.FileExt)
			'判断文件类型
			if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
				CheckFileExt(fileEXT)=false
			end if
			if CheckFileExt(fileEXT)=false then
				response.write "文件格式不正确　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
				exit sub
			end if
			if isnull(myvip) or myvip<>1 then
				curamount=clng(Forum_user(36))*((file.FileSize+1023)\1024)
			else
				curamount=clng(Forum_user(37))*((file.FileSize+1023)\1024)
			end if
			totalamount=totalamount+curamount
			if mymoney<curamount then
				Response.Write "上传需要"&TotalAmount&"元，您的现金不足！ [ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
				exit sub
			end if
			'付值变量
			randomize
			ranNum=int(90000*rnd)+10000
			F_Type=CheckFiletype(fileEXT)
			file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
			filename=file_name&"."&fileExt
			rename=filename&"|"
			filename=formPath&filename
			Filesize=file.FileSize
			'记录文件
			if Filesize>0 then         '如果 FileSize > 0 说明有文件数据
				file.SaveToFile Server.mappath(FileName)   ''执行上传文件
				'创建生成预览图片
				if upload_view=1 and F_Type=1 then
					F_Viewname=previewpath&"pre"&file_name&".JPG"
					call CreateView(FileName,file_name,previewpath,120)
				end if
				call checksave()			'记录文件
				conn.execute("update [user] set userwealth=userwealth-"&curamount&" where username='"&membername&"'")
			end if
			set file=nothing
		next
	end if
	set upload=nothing
	if upNum < int(GroupSetting(40)) and dateupnum < clng(GroupSetting(50)) then
		response.write upNum&"个文件上传成功，共花费"&totalamount&"元！ [ <a href="&Request.ServerVariables("HTTP_REFERER")&">继续上传</a> ]"
	else
		response.write upNum&"个文件上传成功，共花费"&totalamount&"元！本次已达到上传数上限。"
	end if
end sub

''===========================lyfupload组件上传============================
sub upload_1()
	dim obj,filepath,fileExt_a
	dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'限制大小
	obj.maxsize=int(GroupSetting(44))*1024
	'限制类型
	obj.extname=replace(Board_Setting(19),"|",",")
	formPath=obj.request("filepath")
	'在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	randomize
	ranNum=int(90000*rnd)+10000
	filepath=Server.MapPath(formPath)
	if obj.request("fname")="" or isnull(obj.request("fname")) then
		Response.Write "请先选择你要上传的文件　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
		exit sub
	end if
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=file_name&"."&fileExt
	rename=filename & "|"
	'保存文件
	ss=obj.SaveFile("file1",filepath, true,filename)
	dim totalsize,totalamount
	totalsize=obj.FileSize
	if isnull(myvip) or myvip<>1 then
		totalamount=clng(Forum_user(36))*((totalsize+1023)\1024)
	else
		totalamount=clng(Forum_user(37))*((totalsize+1023)\1024)
	end if
	if mymoney<TotalAmount then
		Response.Write "上传需要"&TotalAmount&"元，您的现金不足 [ <a href="&Request.ServerVariables("HTTP_REFERER")&")>重新上传</a> ]"
		exit sub
	end if
	if ss= "3" then
		Response.Write ("文件名重复!　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]")
		exit sub
	elseif ss= "0" then
		Response.Write ("文件大小超过了限制 "&GroupSetting(44)&"K　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]")
		exit sub
	elseif ss = "1" then
		Response.Write ("文件不是指定类型文件!　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]")
		exit sub
	elseif ss = "" then
		Response.Write ("文件上传失败!　[ <a href="&Request.ServerVariables("HTTP_REFERER")&")>重新上传</a> ]")
		exit sub
	else
		'判断文件类型及付值变量
		F_Type=CheckFiletype(fileEXT)
		filename=formPath&filename
		Filesize=obj.filesize
		'创建生成预览图片
		if upload_view=1 and F_Type=1 then
			F_Viewname=previewpath&"pre"&file_name&".JPG"
			call CreateView(FileName,file_name,previewpath,120)
		end if
		'记录文件
		call checksave()			'记录文件
	end if
	set obj=nothing
	conn.execute("update [user] set userwealth=userwealth-"&totalamount&" where username='"&membername&"'")
	if upNum < int(GroupSetting(40)) and dateupnum < clng(GroupSetting(50)) then
		response.write upNum&"个文件上传成功，共花费"&totalamount&"元！ [ <a href="&Request.ServerVariables("HTTP_REFERER")&">继续上传</a> ]"
	else
		response.write upNum&"个文件上传成功，共花费"&totalamount&"元！本次已达到上传数上限。"
	end if
end sub

Private sub checksave()
	'插入上传表并获得ID
	if upload_view=1 and F_Type=1 then
		conn.execute("insert into dv_upfile (F_BoardID,F_UserID,F_Username,F_Filename,F_Viewname,F_FileType,F_Type,F_FileSize) values ("&BoardID&","&UserID&",'"&membername&"','"&replace(rename,"|","")&"','"&F_Viewname&"','"&replace(fileExt,".","")&"',"&F_Type&","&Filesize&")")
	else
		conn.execute("insert into dv_upfile (F_BoardID,F_UserID,F_Username,F_Filename,F_FileType,F_Type,F_FileSize) values ("&BoardID&","&UserID&",'"&membername&"','"&replace(rename,"|","")&"','"&replace(fileExt,".","")&"',"&F_Type&","&Filesize&")")
	end if
	set rs=conn.execute("select top 1 F_ID from dv_upfile order by F_ID desc")
	DownloadID=rs(0)
	rename=DownloadID & ","
	set rs=nothing
	if F_Type=1 then
	 	response.write "<script>parent.frmAnnounce.Content.value+='[upload="&fileExt&"]"&filename&"[/upload]'</script>"
	else
	 	response.write "<script>parent.frmAnnounce.Content.value+='[upload="&fileExt&"]viewfile.asp?ID="&DownloadID&"[/upload]'</script>"
	end if
	response.write "<script>parent.frmAnnounce.upfilerename.value+='"&rename&"'</script>"
	upNum=upNum+1
	response.cookies("upNum")=upNum
	dateupnum=dateupnum+1
	response.Cookies("dateupnum").Expires=Date+1
	response.cookies("dateupnum")=dateupnum
end sub

'判断文件类型是否合格
Private Function CheckFileExt (fileEXT)
	dim Forumupload
	Forumupload=split(Board_Setting(19),"|")
	for i=0 to ubound(Forumupload)
		if lcase(fileEXT)=lcase(trim(Forumupload(i))) then
			CheckFileExt=true
			exit Function
		else
			CheckFileExt=false
		end if
	next
End Function

'判断文件类型:0=其它,1=图片,2=FLASH,3=音乐,4=电影
Private Function CheckFiletype(fileEXT)
	dim upFiletype
	dim FilePic,FileVedio,FileSoft,FileFlash,FileMusic
	fileEXT=lcase(replace(fileExt,".",""))
	FilePic=".gif.jpg.jpeg.png.bmp.tif.iff"
	upFiletype=split(FilePic,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=1
			exit Function
		end if
	next
	FileFlash=".swf.swi"
	upFiletype=split(FileFlash,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=2
			exit Function
		end if
	next
	FileMusic=".mid.wav.mp3.rmi.cda"
	upFiletype=split(FileMusic,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=3
			exit Function
		end if
	next
	FileVedio=".avi.mpg.mpeg.ra.ram.wov.asf"
	upFiletype=split(FileVedio,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=4
			exit Function
		end if
	next
	FileSoft=".rar.zip.exe.php.php3.asp.aspx.htm.html.shtml.js.jsp.pdf.inc.doc.txt.chm.hlp"
	CheckFiletype=0
end function

'创建预览图片:call CreateView(原始文件的路径,预览文件名,预览存放的目录,预览图高度)
sub CreateView(imagename,tempfilename,PreviewFolderName,SetPreviewImageSize)
	'定义变量
	dim PreviewImageFolderName
	dim ogvbox
	PreviewImageFolderName = Server.MapPath(PreviewFolderName)
	set ogvbox = Server.CreateObject("CreatePreviewImage.cGvbox")
	ogvbox.SetSavePreviewImagePath=PreviewImageFolderName  '预览图存放路径
	ogvbox.SetPreviewImageSize =SetPreviewImageSize   '预览图高度
	ogvbox.SetImagefile = trim(Server.MapPath(imagename))              'imagename原始文件的物理路径
	'创建预览图的文件
	If ogvbox.DoImageProcess Then
		'dim PreviewimageFilename
		'PreviewimageFilename = "PreviewImage/" & "pre" &  tempfilename &".jpg"
		'生成预览图的物理文件路径:
		'Response.Write  ogvbox.GetPreviewImagefile
		'生成预览图的URL:
		'Response.Write "<a href=" & PreviewimageFilename & " target=_blank> 查看预览图 </a>"
	Else
		'生成预览图错误
		Response.Write "生成预览图错误:"& ogvbox.GetErrString
	End If
	set ogvbox =nothing
end sub%>