<!--#include file="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="z_down_conn.asp"-->
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
<table width=100% border=0  cellspacing="0" cellpadding="0">
	<tr>
		<td class=tablebody1 width=100% height=26 valign=middle><%
			dim upload_type
			dim upload,file,formName,formPath,filename,fileExt
			dim ranNum
			'�ϴ�������0�������
			upload_type=0
			set rs=conndown.execute("select * from [showpage]")
			select case upload_type
			case 0
				call upload_0()
			case else
				response.write "��ϵͳδ���Ų������"
			end select
			rs.close
			set rs=nothing
		%></td>
	</tr>
</table>
</body>
</html>

<%'������ϴ�
sub upload_0()
	set upload=new Upload_file '�����ϴ�����
	upload.GetData(-1)   'ȡ���ϴ�����,���޴�С
	formPath=upload.form("filepath")
	'��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	for each formName in upload.file '�г������ϴ��˵��ļ�
		set file=upload.file(formName)  '����һ���ļ�����
		if file.filesize>rs("uppicsizemax")*1024 then
			response.write "ͼƬ��С�������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			exit sub
		end if
		if file.filesize<rs("uppicsizemin")*1024 then
			response.write "ͼƬ��С�������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			exit sub
		end if
		fileExt=lcase(file.FileExt)
		if instr(1,"|"&rs("uppictype")&"|","|"&fileExt&"|")<=0 then
			response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			exit sub
		end if 
		randomize
		ranNum=int(90000*rnd)+10000
		filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileEXT
		if file.FileSize>0 then         '��� FileSize > 0 ˵�����ļ�����
			file.SaveToFile Server.mappath(filename)   ''�����ļ�
			response.write "<script>parent.document.forms[0].pic.value='"&filename&"'</script>"
		end if
		set file=nothing
	next
	set upload=nothing
	Htmend "ͼƬ�ϴ��ɹ�!"
end sub

sub HtmEnd(Msg)
	response.write msg
end sub%>
