<!--#include file="conn.asp"-->
<!--#include FILE="z_down_conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="z_down_function.asp"-->
<!--#include FILE="upload.inc"--> 
<html>
<head>
<title>�ļ��ϴ�</title>
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
			Server.ScriptTimeOut=999999'Ҫ�������̳֧���ϴ����ļ��Ƚϴ󣬾ͱ������á�
			'�ϴ���ʽupload_typeֵ: 0���������1��lyfupload
			dim upload_type
			upload_type=upmod
			'�������
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
				response.write "��ϵͳδ�����ϴ�����"
			end select
		%></td>
	</tr>
</table>
</body>
</html>

<%'=====================������ϴ�=======================================================
sub upload_0()
	dim upload,file
	set upload=new Upload_file ''�����ϴ�����
	upload.GetData(-1)   'ȡ���ϴ�����,���޴�С
	formPath=upload.form("filepath")
	'��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/"
	for each formName in upload.file ''�г������ϴ��˵��ļ�
		set file=upload.file(formName)  ''����һ���ļ�����
		if file.filesize>rs("upsizemax")*1024 then
			response.write "�ļ���С�������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			exit sub
		end if
		if file.filesize<rs("upsizemin")*1024 then
			response.write "�ļ���С�������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			exit sub
		end if
		'�ж��ļ�����
		fileExt=lcase(file.FileExt)
		if instr(1,"|"&rs("uptype")&"|","|"&fileExt&"|")<=0 then
			response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			exit sub
		end if 
		'��ֵ����
		randomize
		ranNum=int(90000*rnd)+10000
		file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
		file_name=file_name&"."&fileExt
		filename=formPath&file_name
		Filesize=file.FileSize
		'��¼�ļ�
		if Filesize>0 then         '��� FileSize > 0 ˵�����ļ�����
			file.SaveToFile Server.mappath(FileName)   ''ִ���ϴ��ļ�
			response.write "<script>parent.document.forms[0].txtfilename.value='"&file_name&"'</script>"
			response.write "<script>parent.document.forms[0].localdown.checked=true</script>"
		end if
		set file=nothing
	next
	set upload=nothing
	Htmend "�ļ��ϴ��ɹ�!"
end sub
'=============������ϴ�����====================================

'===========Lyfupload����ϴ�(upload_1)=========================
sub upload_1()
	dim obj,filepath,filename,fs,fileExt
	dim fileExt_a
	dim ss,aa1,aa
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
	obj.maxsize=rs("upsizemax")
	'����
	obj.extname=replace(rs("uptype"),"|",",")
	'������
	formPath=obj.request("filepath")
	'��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	randomize
	ranNum=int(90000*rnd)+10000
	filepath=Server.MapPath(formPath)
	if obj.request("fname")="" or isnull(obj.request("fname")) then
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		exit sub
	end if
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=file_name&"."&fileExt
	ss=obj.SaveFile("file1",filepath, true,filename)
	if ss= "3" then
		Response.Write ("�ļ����ظ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		exit sub
	elseif ss= "0" then
		Response.Write ("�ļ��ߴ����!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		exit sub
	elseif ss = "1" then
		Response.Write ("�ļ�����ָ�������ļ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a>  ]")
		exit sub
	elseif ss = "" then
		Response.Write ("�ļ��ϴ�ʧ��!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		exit sub
	else
		response.write "<script>parent.document.forms[0].txtfilename.value='" & filename & "'</script>"
		response.write "<script>parent.document.forms[0].localdown.checked=true</script>"
		Response.Write "�ļ��ϴ��ɹ���" 
	end if
	set obj=nothing
end sub
'===============Lyfupload����ϴ�(upload_1)����===============================
sub HtmEnd(Msg)
	response.write "�ļ��ϴ��ɹ�,��copy�±ߵ��ļ�����,�Ա�����"
end sub%>
