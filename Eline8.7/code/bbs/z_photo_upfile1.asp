<!--#include FILE="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!--#include FILE="upload.inc"-->
<!-- #include file="inc/const.asp" -->
<%response.buffer=true%>
<html>
<head>
<title>�ļ��ϴ�</title>
<!--#include file="inc/Forum_css.asp"-->
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" border=0 cellspacing=0 cellpadding=0>
	<tr>
		<td class=tablebody1 valign=top height=40><%
			'�ϴ�������0��������ϴ���1��lyfupload
			dim upload_type,up1
			upload_type=0			'�ϴ�������0���������1��Lyfupload
			'----------------------������������----------------
			dim z_photo_t_size,z_photo_size,z_photo_type
			set rs=server.createobject("adodb.recordset")
			sql="select * from z_photo_config where id=1"
			rs.open sql,z_photo_cn,1,1
			z_photo_t_size=rs("z_photo_totle_size")
			z_photo_size=rs("z_photo_size")
			z_photo_type=rs("z_photo_type")
			rs.close
			dim upload,file,formName,formPath,iCount,filename,fileExt
			dim upNum
			dim uploadsuc
			dim Forumupload
			dim ranNum
			dim uploadfiletype
			dim pho_title,pho_brief,pho_size,tsize
			dim tshared,tpho_groups
			upNum=request.cookies("upNum")
			if not isnumeric(upnum) then upnum=1
			if Cint(GroupSetting(7))=0 then
				response.write "��û���ڱ���̳�ϴ��ļ���Ȩ��"
			elseif not founduser then
				response.write "ֻ�е�¼�û������ϴ���"
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
<%
'������ϴ�
sub upload_0()
	set upload=new Upload_file
	upload.GetData (z_photo_size)
	if upload.err > 0 then  '�������
		select case upload.err
		case 1
			Response.Write "����ѡ����Ҫ�ϴ����ļ���"
		case 2
			Response.Write "ͼƬ��С���������� "&z_photo_size&" ��"
		end select
	else
		dim curamount,totalamount
		formPath=upload.form("filepath")
		if right(formPath,1)<>"/" then formPath=formPath&"/" 
		for each formName in upload.file
			set file=upload.file(formName)
			pho_title=upload.form("pho_title")
			tpho_groups=upload.form("pho_groups")
			pho_brief=upload.form("pho_brief")
			tpho_groups=cint(tpho_groups)
 			set rs=server.createobject("adodb.recordset")
			sql="select SUM(pho_group_size) as ts from z_photo_group where username='"&membername&"' and pho_state='v'"
			rs.open sql,z_photo_cn,1,3
			tsize=rs("ts")+ file.filesize
			rs.close
			if tsize>z_photo_t_size then
				response.write "<font size=2>�������ļ���С������������ "&z_photo_t_size&" ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
				exit sub
			elseif file.filesize> z_photo_size then
				response.write "<font size=2>�����ļ���С���������� "&z_photo_size&" ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
				exit sub
			end if
			Server.ScriptTimeOut=3*int(file.filesize/1000)
			fileExt=lcase(file.FileExt)
			uploadsuc=false
			Forum_upload=lcase(z_photo_type) '��������ϴ��ĸ�ʽ��
			Forumupload=split(Forum_upload,",")
			for i=0 to ubound(Forumupload)
				if fileEXT=lcase(trim(Forumupload(i))) then
					uploadsuc=true
					exit for
				else
					uploadsuc=false
				end if
			next
			if fileEXT="asp" and fileEXT="asa" and fileEXT="aspx" then
				uploadsuc=false
			end if
			if uploadsuc=false then
				response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
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
			filename=replace(filename,".tmp","")
			filename=replace(filename,"rad","")
			randomize
			ranNum=int(90000*rnd)+10000
			filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
				pho_size=file.FileSize
				file.SaveToFile Server.mappath(FileName)   ''�����ļ�
				conn.execute("update [user] set userwealth=userwealth-"&curamount&" where username='"&membername&"'")
			end if
			set file=nothing
		next
		set upload=nothing
		z_photo_cn.execute("insert into z_photo (username,pho_url,pho_title,pho_brief,pho_group,pho_size) values ('"&membername&"','"&filename&"','"&pho_title&"','"&pho_brief&"','"&tpho_groups&"','"&pho_size&"')")
		call up_group(tpho_groups,pho_size)
		call Close_z_photo_cn()
		call Htmend(TotalAmount)
	end if
end sub

'===========Lyfupload����ϴ�(upload_1)=========================
sub upload_1()
	dim obj,forumPath,filename,fileExt_a
	dim ss,aa1,aa
	dim pho_title,tpho_groups,pho_brief
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
	obj.maxsize=z_photo_size
	'����
	obj.extname=z_photo_type
	'������
	formPath=obj.request("filepath")
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	pho_title=obj.request("pho_title")
	tpho_groups=obj.request("pho_groups")
	pho_brief=obj.request("pho_brief")
	tpho_groups=cint(tpho_groups)
	randomize
	ranNum=int(90000*rnd)+10000
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=filename&"."&fileExt
	ss=obj.SaveFile("file1",Server.MapPath(formPath),true,filename)
	set rs=server.createobject("adodb.recordset")
	pho_size=obj.FileSize   
 	sql="select SUM(pho_group_size) as ts from z_photo_group where username='"&membername&"' and pho_state='v'"
	rs.open sql,z_photo_cn,1,3
	tsize=rs("ts")+ pho_size
	rs.close
	if tsize>z_photo_t_size then
	 	response.write "<font size=2>�������ļ���С������������ "&z_photo_t_size&" ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
		exit sub
	end if
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
		Response.Write ("�ļ����ظ�!")
		exit sub
	elseif ss= "0" then
   	Response.Write ("�ļ��ߴ����!")
		exit sub
	elseif ss = "1" then
 		Response.Write ("�ļ�����ָ�������ļ�!")
		exit sub
	elseif ss = "" then
 		Response.Write ("�ļ��ϴ�ʧ��!")
		exit sub
	end if
	z_photo_cn.execute("insert into z_photo (username,pho_url,pho_title,pho_brief,pho_group,pho_size) values ('"&membername&"','"&filename&"','"&pho_title&"','"&pho_brief&"','"&tpho_groups&"','"&pho_size&"')")
	call up_group(tpho_groups,pho_size)
	call Close_z_photo_cn()
	call HtmEnd()
	set obj=nothing
end sub

'------------------------------
sub HtmEnd(amount)
	response.write "<font size=2>���ļ��ϴ��ɹ���������"&amount&"Ԫ�� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
end sub

'-----------------�������ͳ����Ϣ-����-------------------
sub up_group(tpho_groups,pho_size)
 set rs=server.createobject("adodb.recordset")
  sql="select pho_group_totle,pho_group_size from z_photo_group where pho_group_id = "&tpho_groups
	rs.open sql,z_photo_cn,1,3
	rs("pho_group_totle")=rs("pho_group_totle")+1
	rs("pho_group_size")=rs("pho_group_size")+pho_size
	rs.update
	rs.close	
	Response.Write "<script>"
	Response.Write "alert('���Ѿ��ɹ��ϴ�һ����Ƭ��"&pho_size&"');"
	Response.Write "</script>"
end sub
%>
