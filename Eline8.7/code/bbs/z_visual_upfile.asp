<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="upload.inc"-->
<!--#include file="z_visual_const.asp"-->
<html>
<head>
<title>�ļ��ϴ�</title>
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
			upload_type=0	'�ϴ�������0���������1��Lyfupload
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
				response.write "��ϵͳδ�����ϴ�����"
			end select
		%></td>
	</tr>
</table>
</body>
</html>

<%'===========������ϴ�(upload_0)====================
sub upload_0()
	set upload=new Upload_file ''�����ϴ�����
	upload.GetData (int(PhotoBgSize*1024))   'ȡ���ϴ�����,���޴�С
	if upload.err > 0 then
  	select case upload.err
		case 1
			Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
		case 2
			Response.Write "ͼƬ��С���������� "&PhotoBgSize&"K��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
		end select
		exit sub
	else
		dim curamount,totalamount
		formPath=upload.form("filepath")
		if right(formPath,1)<>"/" then formPath=formPath&"/" 
		for each formName in upload.file ''�г������ϴ��˵��ļ�
			set file=upload.file(formName)  ''����һ���ļ�����
			fileExt=lcase(file.FileExt)
			if CheckFileExt(fileEXT)=false then
				response.write "�ļ���ʽ����ȷ��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
				exit sub
			end if
			curamount=clng(PhotoUpPrice)*((file.FileSize+1023)\1024)
			TotalAmount=TotalAmount+curamount
			if mymoney<curamount then
				Response.Write "�ϴ���Ҫ"&TotalAmount&"Ԫ�������ֽ��� [ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
				exit sub
			end if
			randomize
			ranNum=int(90000*rnd)+10000
			filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
				file.SaveToFile Server.mappath(formPath&filename)   ''�����ļ�
				conn.execute("update [user] set userwealth=userwealth-"&curamount&" where username='"&membername&"'")
				response.write "<script language=javascript>parent.document.all.photo_request.bgfname.value='"&FileName&"';parent.document.all.photo_request.photobg.selectedIndex=2;parent.document.all.photo_request.photobg.fireEvent(""onchange"");</script>"
			end if
			set file=nothing
		next
		call Htmend(TotalAmount)
	end if
	set upload=nothing
end sub

'===========Lyfupload����ϴ�(upload_1)=========================
sub upload_1()
	dim obj,filename,fileExt_a
	dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
	obj.maxsize=int(PhotoUpPrice)*1024
	'����
	obj.extname="gif,jpg,jpeg"
	'������
	formPath=obj.request("filepath")
	'��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	if obj.request("fname")="" or isnull(obj.request("fname")) then
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
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
		Response.Write "�ϴ���Ҫ"&TotalAmount&"Ԫ�������ֽ��� [ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]"
		exit sub
	end if

	if ss= "3" then
 		Response.Write ("�ļ����ظ�![ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]")
	elseif ss= "0" then
 		Response.Write ("�ļ��ߴ����![ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]")
	elseif ss = "1" then
 		Response.Write ("�ļ�����ָ�������ļ�![ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]")
	elseif ss = "" then
 		Response.Write ("�ļ��ϴ�ʧ��![ <a href="&Request.ServerVariables("HTTP_REFERER")&">�����ϴ�</a> ]")
	else
		response.write "<script language=javascript>parent.document.all.photo_request.bgfname.value='" &formPath&filename & "'</script>"
		call Htmend(TotalAmount)
	end if
	set obj=nothing
	conn.execute("update [user] set userwealth=userwealth-"&totalamount&" where username='"&membername&"'")
end sub

sub HtmEnd(amount)
 response.write "ͼƬ�ϴ��ɹ���������"&amount&"Ԫ��[ <a href="&Request.ServerVariables("HTTP_REFERER")&">����</a> ]"
end sub

'�ж��ļ������Ƿ�ϸ�
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