<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="upload.inc"-->
<!--#include file="z_visual_const.asp"-->
<html>
<head>
<title>文件上传</title>
<!--#include file=inc/Forum_css.asp-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body topmargin="0" leftmargin="0">
<script>
parent.document.all.photo_request.Submit.disabled=false;
parent.document.all.photo_request.Clear.disabled=false;
</script>
<table width=100% border=0  cellspacing="0" cellpadding="0">
	<tr>
		<td class=tablebody2 width=100% height=26 valign=middle><%
			dim upload_type
			dim upload,file,formName,formPath,filename,fileExt
			dim ranNum
			upload_type=0	'上传方法，0＝无组件，1＝Lyfupload
			dim filepaths,objFSO
			dim rsVisualConfig,PhotoBgSize,PhotoUpPrice
			set rsVisualConfig=server.createobject("ADODB.Recordset")
			rsVisualConfig.open "select PhotoBgSize,PhotoUpPrice1,PhotoUpPrice2 from visual_Config",conn,1,1
			PhotoBgSize=rsVisualConfig(0)
			if isnull(v_myvip) or v_myvip<>1 then
				PhotoUpPrice=rsVisualConfig(1)
			else
				PhotoUpPrice=rsVisualConfig(2)
			end if
			set rsVisualConfig=nothing
			if isnull(request("oldfname")) or request("oldfname")="" then
				filepaths=""
			else
				filepaths=checkstr(trim(request("oldfname")))
			end if
			if filepaths<>"" then
				on error resume next
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				set rs=conn.execute("select background from visual_photo where background='"&filepaths&"'")
				if rs.eof then
					filepaths=Server.MapPath(PhotoPath&"\"&filepaths)
					if objFSO.fileExists(filepaths) then
						objFSO.DeleteFile(filepaths)
					end if
				end if
				set rs=nothing
			end if
			
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

<%'===========无组件上传(upload_0)====================
sub upload_0()
	set upload=new Upload_file ''建立上传对象
	upload.GetData (int(PhotoBgSize*1024))   '取得上传数据,不限大小
	if upload.err > 0 then
  	select case upload.err
		case 1
			Response.Write "请先选择你要上传的文件　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
		case 2
			Response.Write "图片大小超过了限制 "&PhotoBgSize&"K　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
		end select
		exit sub
	else
		dim curamount,totalamount
		formPath=upload.form("filepath")
		if right(formPath,1)<>"/" then formPath=formPath&"/" 
		for each formName in upload.file ''列出所有上传了的文件
			set file=upload.file(formName)  ''生成一个文件对象
			fileExt=lcase(file.FileExt)
			if CheckFileExt(fileEXT)=false then
				response.write "文件格式不正确　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
				exit sub
			end if
			curamount=clng(PhotoUpPrice)*((file.FileSize+1023)\1024)
			TotalAmount=TotalAmount+curamount
			if mymoney<curamount then
				Response.Write "上传需要"&TotalAmount&"元，您的现金不足 [ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
				exit sub
			end if
			randomize
			ranNum=int(90000*rnd)+10000
			filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
				file.SaveToFile Server.mappath(formPath&filename)   ''保存文件
				conn.execute("update [user] set userwealth=userwealth-"&curamount&" where username='"&membername&"'")
				response.write "<script language=javascript>parent.document.all.photo_request.bgfname.value='"&FileName&"';parent.document.all.photo_request.photobg.selectedIndex=2;parent.document.all.photo_request.photobg.fireEvent(""onchange"");</script>"
			end if
			set file=nothing
		next
		call Htmend(TotalAmount)
	end if
	set upload=nothing
end sub

'===========Lyfupload组件上传(upload_1)=========================
sub upload_1()
	dim obj,filename,fileExt_a
	dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'大小
	obj.maxsize=int(PhotoUpPrice)*1024
	'类型
	obj.extname="gif,jpg,jpeg"
	'重命名
	formPath=obj.request("filepath")
	'在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	if obj.request("fname")="" or isnull(obj.request("fname")) then
		Response.Write "请先选择你要上传的文件　[ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
		exit sub
	end if
	randomize
	ranNum=int(90000*rnd)+10000
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=filename&"."&fileExt
	ss=obj.SaveFile("file1",Server.MapPath(formPath), true,filename)
	dim totalsize,totalamount
	totalsize=obj.FileSize
	totalamount=clng(PhotoUpPrice)*((totalsize+1023)\1024)
	if mymoney<TotalAmount then
		Response.Write "上传需要"&TotalAmount&"元，您的现金不足 [ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]"
		exit sub
	end if

	if ss= "3" then
 		Response.Write ("文件名重复![ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]")
	elseif ss= "0" then
 		Response.Write ("文件尺寸过大![ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]")
	elseif ss = "1" then
 		Response.Write ("文件不是指定类型文件![ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]")
	elseif ss = "" then
 		Response.Write ("文件上传失败![ <a href="&Request.ServerVariables("HTTP_REFERER")&">重新上传</a> ]")
	else
		response.write "<script language=javascript>parent.document.all.photo_request.bgfname.value='" &formPath&filename & "'</script>"
		call Htmend(TotalAmount)
	end if
	set obj=nothing
	conn.execute("update [user] set userwealth=userwealth-"&totalamount&" where username='"&membername&"'")
end sub

sub HtmEnd(amount)
 response.write "图片上传成功，共花费"&amount&"元！[ <a href="&Request.ServerVariables("HTTP_REFERER")&">返回</a> ]"
end sub

'判断文件类型是否合格
Private Function CheckFileExt (fileEXT)
dim Forumupload
Forumupload="gif,jpg,jpeg"
Forumupload=split(Forumupload,",")
	for i=0 to ubound(Forumupload)
		if lcase(fileEXT)=lcase(trim(Forumupload(i))) then
			CheckFileExt=true
			exit Function
		else
			CheckFileExt=false
		end if
	next
End Function%>