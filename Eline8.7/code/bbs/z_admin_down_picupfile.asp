<!--#include file="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="z_down_conn.asp"-->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>文件上传</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body topmargin="0" leftmargin="0">
<script>
parent.document.all.frmAnnounce.cmdok.disabled=false;
parent.document.all.frmAnnounce.cmdcancel.disabled=false;
</script>
<table width=100% border=0  cellspacing="0" cellpadding="0">
	<tr>
		<td class=forumrow width=100% height=26 valign=middle><%
			dim upload_type
			dim upload,file,formName,formPath,filename,fileExt
			dim ranNum
			'上传方法，0＝无组件
			upload_type=0
			set rs=conndown.execute("select * from [showpage]")
			select case upload_type
			case 0
				call upload_0()
			case else
				response.write "本系统未开放插件功能"
			end select
			rs.close
			set rs=nothing
		%></td>
	</tr>
</table>
</body>
</html>

<%'无组件上传
sub upload_0()
	set upload=new Upload_file '建立上传对象
	upload.GetData(-1)   '取得上传数据,不限大小
	formPath=upload.form("filepath")
	'在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	for each formName in upload.file '列出所有上传了的文件
		set file=upload.file(formName)  '生成一个文件对象
		if file.filesize>rs("uppicsizemax")*1024 then
			response.write "图片大小超过限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			exit sub
		end if
		if file.filesize<rs("uppicsizemin")*1024 then
			response.write "图片大小低于限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			exit sub
		end if
		fileExt=lcase(file.FileExt)
		if instr(1,"|"&rs("uppictype")&"|","|"&fileExt&"|")<=0 then
			response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			exit sub
		end if 
		randomize
		ranNum=int(90000*rnd)+10000
		filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileEXT
		if file.FileSize>0 then         '如果 FileSize > 0 说明有文件数据
			file.SaveToFile Server.mappath(filename)   ''保存文件
			response.write "<script>parent.document.forms[0].pic.value='"&filename&"'</script>"
		end if
		set file=nothing
	next
	set upload=nothing
	Htmend "图片上传成功!"
end sub

sub HtmEnd(Msg)
	response.write msg
end sub%>
