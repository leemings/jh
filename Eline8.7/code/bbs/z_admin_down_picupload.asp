<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<html>
<head>
<title></title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width=100%>
	<tr>
		<td class=forumrow valign=middle height=24>
			<form name="form" method="post" action="z_admin_down_picupfile.asp" enctype="multipart/form-data" >
			<input type="hidden" name="filepath" value="uploadimages/">
			<input type="hidden" name="act" value="upload">
			<input type="file" name="file1">
			<input type="submit" name="Submit" value="上传" onclick="parent.document.all.frmAnnounce.cmdok.disabled=true;parent.document.all.frmAnnounce.cmdcancel.disabled=true;">
			</form>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close
set conn=nothing%>