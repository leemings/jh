<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>�ļ��ϴ�</title>
<!--#include file=inc/Forum_css.asp-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body topmargin="0" leftmargin="0">
<script>
parent.document.all.theForm.Submit.disabled=false;
parent.document.all.theForm.Submit2.disabled=false;
</script>
<table width=100% border=0  cellspacing="0" cellpadding="0">
	<tr>
		<td class=tablebody1 width=100% height=26 valign=middle><%
			dim upload_type
			dim upload,file,formName,formPath,filename,fileExt
			dim ranNum
			upload_type=0	'�ϴ�������0���������1��Lyfupload
			
			'******************
			'ɾ���û��ɵ�ͷ��
			dim filepaths,objFSO,upface
			if  founduser then
				on error resume next
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				set rs=conn.execute("select face from [user] where userid="&userid)
		    upface=rs(0)
				if instr(upface,"uploadFace")>0  then
					filepaths=Server.MapPath(""&upface&"")
					if objFSO.fileExists(filepaths) then
						objFSO.DeleteFile(filepaths)
					end if
				end if
				set rs=nothing
			end if
			'******************
			'ɾ���û��ɵ�ͷ��
			
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
	upload.GetData (int(Forum_Setting(56))*1024)   'ȡ���ϴ�����,���޴�С
	if upload.err > 0 then
  	select case upload.err
		case 1
			Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		case 2
			Response.Write "ͼƬ��С���������� "&Forum_Setting(56)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
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
				response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				exit sub
			end if
			if isnull(myvip) or myvip<>1 then
				curamount=clng(Forum_user(34))*((file.FileSize+1023)\1024)
			else
				curamount=clng(Forum_user(35))*((file.FileSize+1023)\1024)
			end if
			TotalAmount=TotalAmount+curamount
			if mymoney<curamount then
				Response.Write "�ϴ���Ҫ"&TotalAmount&"Ԫ�������ֽ��� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				exit sub
			end if
			randomize
			ranNum=int(90000*rnd)+10000
			filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
				file.SaveToFile Server.mappath(filename)   ''�����ļ�
				conn.execute("update [user] set userwealth=userwealth-"&curamount&" where username='"&membername&"'")
				response.write "<script>parent.document.all.theForm.myface.value='"&FileName&"'</script>"
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
	obj.maxsize=int(Forum_Setting(56))*1024
	'����
	obj.extname="gif,jpg,bmp,jpeg"
	'������
	formPath=obj.request("filepath")
	'��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	if obj.request("fname")="" or isnull(obj.request("fname")) then
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
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
	if isnull(myvip) or myvip<>1 then
		totalamount=clng(Forum_user(34))*((totalsize+1023)\1024)
	else
		totalamount=clng(Forum_user(35))*((totalsize+1023)\1024)
	end if
	if mymoney<TotalAmount then
		Response.Write "�ϴ���Ҫ"&TotalAmount&"Ԫ�������ֽ��� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		exit sub
	end if

	if ss= "3" then
 		Response.Write ("�ļ����ظ�![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
	elseif ss= "0" then
 		Response.Write ("�ļ��ߴ����![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
	elseif ss = "1" then
 		Response.Write ("�ļ�����ָ�������ļ�![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
	elseif ss = "" then
 		Response.Write ("�ļ��ϴ�ʧ��![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
	else
		response.write "<script>parent.document.all.theForm.myface.value='" &formPath&filename & "'</script>"
		call Htmend(TotalAmount)
	end if
	set obj=nothing
	conn.execute("update [user] set userwealth=userwealth-"&totalamount&" where username='"&membername&"'")
end sub

sub HtmEnd(amount)
 response.write "ͼƬ�ϴ��ɹ���������"&amount&"Ԫ��"
end sub

'�ж��ļ������Ƿ�ϸ�
Private Function CheckFileExt (fileEXT)
dim Forumupload
Forumupload="gif,jpg,bmp,jpeg"
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