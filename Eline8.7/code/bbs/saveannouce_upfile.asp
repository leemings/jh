<!--#include FILE="conn.asp"-->
<!--#include FILE="upload.inc"-->
<!--#include file="inc/const.asp" -->
<html>
<head>
<title>�ļ��ϴ�</title>
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
			Server.ScriptTimeOut=999999'Ҫ�������̳֧���ϴ����ļ��Ƚϴ󣬾ͱ������á�
			'�ϴ���ʽupload_typeֵ: 0���������1��lyfupload
			dim upload_type
			upload_type=0
			'��������Ԥ��ͼƬ,��ҪCreatePreviewImage���֧��,upload_viewֵ: 0=��֧��,1=֧��(��Ŀ¼��Ҫ��PreviewImage�ļ��д���ļ�)
			dim upload_view
			upload_view=0
			'�������
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
				response.write "��û���ڱ���̳�ϴ��ļ���Ȩ��"
			elseif not founduser then
				response.write "ֻ�е�¼�û������ϴ���"
			elseif upNum >= int(GroupSetting(40)) then
				response.write "һ��ֻ���ϴ�"&GroupSetting(40)&"���ļ���"
			elseif dateupnum >= int(GroupSetting(50)) then
				response.write "�������ϴ����ļ��ѳ�����"&GroupSetting(50)&"����"
			else
				select case upload_type
				case 0
					call upload_0()
				case 1
					call upload_1()
				case else
					response.write "��ϵͳδ�����ϴ�����"
				end select
			end if
		%></td>
	</tr>
</table>
</body>
</html>
<%'===========================������ϴ�============================
sub upload_0()
	dim upload,file
	set upload=new Upload_file ''�����ϴ�����
	upload.GetData (GroupSetting(44)*1024)   'ȡ���ϴ�����,���޴�С
	if upload.err > 0 then
		select case upload.err
		case 1
			Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
		case 2
			Response.Write "�ļ���С���������� "&GroupSetting(44)&"K��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
		end select
		exit sub
	else
		dim curamount,totalamount
		formPath=upload.form("filepath")
		'��Ŀ¼���(/)
		if right(formPath,1)<>"/" then formPath=formPath&"/"
		for each formName in upload.file ''�г������ϴ��˵��ļ�
			if upNum >= int(GroupSetting(40)) or dateupnum >= clng(GroupSetting(50)) then
				response.write "�Ѵﵽ�ϴ��������ޡ�"
				exit sub
			end if
			set file=upload.file(formName)  ''����һ���ļ�����
			fileExt=lcase(file.FileExt)
			'�ж��ļ�����
			if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
				CheckFileExt(fileEXT)=false
			end if
			if CheckFileExt(fileEXT)=false then
				response.write "�ļ���ʽ����ȷ��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
				exit sub
			end if
			if isnull(myvip) or myvip<>1 then
				curamount=clng(Forum_user(36))*((file.FileSize+1023)\1024)
			else
				curamount=clng(Forum_user(37))*((file.FileSize+1023)\1024)
			end if
			totalamount=totalamount+curamount
			if mymoney<curamount then
				Response.Write "�ϴ���Ҫ"&TotalAmount&"Ԫ�������ֽ��㣡 [ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
				exit sub
			end if
			'��ֵ����
			randomize
			ranNum=int(90000*rnd)+10000
			F_Type=CheckFiletype(fileEXT)
			file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
			filename=file_name&"."&fileExt
			rename=filename&"|"
			filename=formPath&filename
			Filesize=file.FileSize
			'��¼�ļ�
			if Filesize>0 then         '��� FileSize > 0 ˵�����ļ�����
				file.SaveToFile Server.mappath(FileName)   ''ִ���ϴ��ļ�
				'��������Ԥ��ͼƬ
				if upload_view=1 and F_Type=1 then
					F_Viewname=previewpath&"pre"&file_name&".JPG"
					call CreateView(FileName,file_name,previewpath,120)
				end if
				call checksave()			'��¼�ļ�
				conn.execute("update [user] set userwealth=userwealth-"&curamount&" where username='"&membername&"'")
			end if
			set file=nothing
		next
	end if
	set upload=nothing
	if upNum < int(GroupSetting(40)) and dateupnum < clng(GroupSetting(50)) then
		response.write upNum&"���ļ��ϴ��ɹ���������"&totalamount&"Ԫ�� [ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
	else
		response.write upNum&"���ļ��ϴ��ɹ���������"&totalamount&"Ԫ�������Ѵﵽ�ϴ������ޡ�"
	end if
end sub

''===========================lyfupload����ϴ�============================
sub upload_1()
	dim obj,filepath,fileExt_a
	dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'���ƴ�С
	obj.maxsize=int(GroupSetting(44))*1024
	'��������
	obj.extname=replace(Board_Setting(19),"|",",")
	formPath=obj.request("filepath")
	'��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	randomize
	ranNum=int(90000*rnd)+10000
	filepath=Server.MapPath(formPath)
	if obj.request("fname")="" or isnull(obj.request("fname")) then
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
		exit sub
	end if
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=file_name&"."&fileExt
	rename=filename & "|"
	'�����ļ�
	ss=obj.SaveFile("file1",filepath, true,filename)
	dim totalsize,totalamount
	totalsize=obj.FileSize
	if isnull(myvip) or myvip<>1 then
		totalamount=clng(Forum_user(36))*((totalsize+1023)\1024)
	else
		totalamount=clng(Forum_user(37))*((totalsize+1023)\1024)
	end if
	if mymoney<TotalAmount then
		Response.Write "�ϴ���Ҫ"&TotalAmount&"Ԫ�������ֽ��� [ <a href="&Request.ServerVariables("HTTP_REFERER")&")>�����ϴ�</a> ]"
		exit sub
	end if
	if ss= "3" then
		Response.Write ("�ļ����ظ�!��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]")
		exit sub
	elseif ss= "0" then
		Response.Write ("�ļ���С���������� "&GroupSetting(44)&"K��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]")
		exit sub
	elseif ss = "1" then
		Response.Write ("�ļ�����ָ�������ļ�!��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]")
		exit sub
	elseif ss = "" then
		Response.Write ("�ļ��ϴ�ʧ��!��[ <a href="&Request.ServerVariables("HTTP_REFERER")&")>�����ϴ�</a> ]")
		exit sub
	else
		'�ж��ļ����ͼ���ֵ����
		F_Type=CheckFiletype(fileEXT)
		filename=formPath&filename
		Filesize=obj.filesize
		'��������Ԥ��ͼƬ
		if upload_view=1 and F_Type=1 then
			F_Viewname=previewpath&"pre"&file_name&".JPG"
			call CreateView(FileName,file_name,previewpath,120)
		end if
		'��¼�ļ�
		call checksave()			'��¼�ļ�
	end if
	set obj=nothing
	conn.execute("update [user] set userwealth=userwealth-"&totalamount&" where username='"&membername&"'")
	if upNum < int(GroupSetting(40)) and dateupnum < clng(GroupSetting(50)) then
		response.write upNum&"���ļ��ϴ��ɹ���������"&totalamount&"Ԫ�� [ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
	else
		response.write upNum&"���ļ��ϴ��ɹ���������"&totalamount&"Ԫ�������Ѵﵽ�ϴ������ޡ�"
	end if
end sub

Private sub checksave()
	'�����ϴ������ID
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

'�ж��ļ������Ƿ�ϸ�
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

'�ж��ļ�����:0=����,1=ͼƬ,2=FLASH,3=����,4=��Ӱ
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

'����Ԥ��ͼƬ:call CreateView(ԭʼ�ļ���·��,Ԥ���ļ���,Ԥ����ŵ�Ŀ¼,Ԥ��ͼ�߶�)
sub CreateView(imagename,tempfilename,PreviewFolderName,SetPreviewImageSize)
	'�������
	dim PreviewImageFolderName
	dim ogvbox
	PreviewImageFolderName = Server.MapPath(PreviewFolderName)
	set ogvbox = Server.CreateObject("CreatePreviewImage.cGvbox")
	ogvbox.SetSavePreviewImagePath=PreviewImageFolderName  'Ԥ��ͼ���·��
	ogvbox.SetPreviewImageSize =SetPreviewImageSize   'Ԥ��ͼ�߶�
	ogvbox.SetImagefile = trim(Server.MapPath(imagename))              'imagenameԭʼ�ļ�������·��
	'����Ԥ��ͼ���ļ�
	If ogvbox.DoImageProcess Then
		'dim PreviewimageFilename
		'PreviewimageFilename = "PreviewImage/" & "pre" &  tempfilename &".jpg"
		'����Ԥ��ͼ�������ļ�·��:
		'Response.Write  ogvbox.GetPreviewImagefile
		'����Ԥ��ͼ��URL:
		'Response.Write "<a href=" & PreviewimageFilename & " target=_blank> �鿴Ԥ��ͼ </a>"
	Else
		'����Ԥ��ͼ����
		Response.Write "����Ԥ��ͼ����:"& ogvbox.GetErrString
	End If
	set ogvbox =nothing
end sub%>