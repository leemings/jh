<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_visual_const.asp"-->
<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body leftmargin="0" topmargin="0">
<form name="form" method="post" action="z_visual_upfile.asp" enctype="multipart/form-data" >
<table border="0"  cellspacing="0" cellpadding="0" width=100%>
	<tr>
		<td class=tablebody2 height=26><input type="hidden" name="filepath" value="<%=PhotoPath%>/"><input type="hidden" name="act" value="upload"><input type="file" name="file1"><input type="hidden" name="fname"><input type="submit" name="Submit" value="上传" onclick="form.action='z_visual_upfile.asp?oldfname='+parent.document.all.photo_request.bgfname.value;fname.value=document.all.file1.value;parent.document.all.photo_request.Submit.disabled=true;parent.document.all.photo_request.Clear.disabled=true;"><%
			dim rsVisualConfig
			set rsVisualConfig=server.createobject("ADODB.Recordset")
			rsVisualConfig.open "select PhotoBgSize,PhotoUpPrice1,PhotoUpPrice2 from visual_Config",conn,1,1
			response.write "<font color="&Forum_body(8)&">&nbsp;背景最大"&rsVisualConfig(0)&"K，每K文件"
			if isnull(v_myvip) or v_myvip<>1 then
				response.write rsVisualConfig(1)
			else
				response.write rsVisualConfig(2)
			end if
			response.write "元，现有"&mymoney&"元</font>"
			set rsVisualConfig=nothing
		%></td>
	</tr>
</table>
</form>
</body>
</html>
<%call CloseDatabase%>