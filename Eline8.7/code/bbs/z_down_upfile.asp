<!--#include file="conn.asp"-->
<!--#include FILE="z_down_conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="z_down_function.asp"-->
<!--#include FILE="upload.inc"--> 
<html>
<head>
<title>文件上传</title>
<!--#include file=inc/Forum_css.asp-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body topmargin="0" leftmargin="0">
<script>
parent.document.all.frmAnnounce.cmdok.disabled=false;
parent.document.all.frmAnnounce.cmdcancel.disabled=false;
</script>
<table width="100%" border=0 cellspacing=0 cellpadding=0>
	<tr>
		<td class=tablebody1 height=26 width=100% valign=middle><%
			Server.ScriptTimeOut=999999'要是你的论坛支持上传的文件比较大，就必须设置。
			'上传方式upload_type值: 0＝无组件，1＝lyfupload
			dim upload_type
			upload_type=upmod
			'定义变量
			dim Forumupload,ranNum
			dim formName,formPath,filename,file_name,fileExt,Filesize
			dim upNum,dateupnum
			dim rename,DownloadID
			set rs=conndown.execute("select * from [showpage]")
			select case upload_type
			case 0
				call upload_0()
			case 1
				call upload_1()
			case else
				response.write "本系统未开放上传功能"
			end select
		%></td>
	</tr>
</table>
</body>
</html>

<%'=====================无组件上传=======================================================
sub upload_0()
	dim upload,file
	set upload=new Upload_file ''建立上传对象
	upload.GetData(-1)   '取得上传数据,不限大小
	formPath=upload.form("filepath")
	'在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/"
	for each formName in upload.file ''列出所有上传了的文件
		set file=upload.file(formName)  ''生成一个文件对象
		if file.filesize>rs("upsizemax")*1024 then
			response.write "文件大小超过限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			exit sub
		end if
		if file.filesize<rs("upsizemin")*1024 then
			response.write "文件大小低于限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			exit sub
		end if
		'判断文件类型
		fileExt=lcase(file.FileExt)
		if instr(1,"|"&rs("uptype")&"|","|"&fileExt&"|")<=0 then
			response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			exit sub
		end if 
		'付值变量
		randomize
		ranNum=int(90000*rnd)+10000
		file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
		file_name=file_name&"."&fileExt
		filename=formPath&file_name
		Filesize=file.FileSize
		'记录文件
		if Filesize>0 then         '如果 FileSize > 0 说明有文件数据
			file.SaveToFile Server.mappath(FileName)   ''执行上传文件
			response.write "<script>parent.document.forms[0].txtfilename.value='"&file_name&"'</script>"
			response.write "<script>parent.document.forms[0].localdown.checked=true</script>"
		end if
		set file=nothing
	next
	set upload=nothing
	Htmend "文件上传成功!"
end sub
'=============无组件上传结束====================================

'===========Lyfupload组件上传(upload_1)=========================
sub upload_1()
	dim obj,filepath,filename,fs,fileExt
	dim fileExt_a
	dim ss,aa1,aa
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'大小
	obj.maxsize=rs("upsizemax")
	'类型
	obj.extname=replace(rs("uptype"),"|",",")
	'重命名
	formPath=obj.request("filepath")
	'在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	randomize
	ranNum=int(90000*rnd)+10000
	filepath=Server.MapPath(formPath)
	if obj.request("fname")="" or isnull(obj.request("fname")) then
		Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		exit sub
	end if
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=file_name&"."&fileExt
	ss=obj.SaveFile("file1",filepath, true,filename)
	if ss= "3" then
		Response.Write ("文件名重复!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		exit sub
	elseif ss= "0" then
		Response.Write ("文件尺寸过大!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		exit sub
	elseif ss = "1" then
		Response.Write ("文件不是指定类型文件!　[ <a href=# onclick=history.go(-1)>重新上传</a>  ]")
		exit sub
	elseif ss = "" then
		Response.Write ("文件上传失败!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		exit sub
	else
		response.write "<script>parent.document.forms[0].txtfilename.value='" & filename & "'</script>"
		response.write "<script>parent.document.forms[0].localdown.checked=true</script>"
		Response.Write "文件上传成功！" 
	end if
	set obj=nothing
end sub
'===============Lyfupload组件上传(upload_1)结束===============================
sub HtmEnd(Msg)
	response.write "文件上传成功,请copy下边的文件连接,以备后用"
end sub%>
