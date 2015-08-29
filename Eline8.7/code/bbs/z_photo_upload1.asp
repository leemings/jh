<!--#include file="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!-- #include file="inc/const.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<script language="JavaScript" type="text/JavaScript">
function z_photo_checkform()
{
if(document.all["pho_name12"].value==""){
	alert("照片地址不能为空！");
	return false;
}
if(document.all["pho_title"].value==""){
	alert("照片标题不能为空！");
	return false;
}
if(document.all["pho_brief"].value==""){
	alert("照片说明不能为空！");
	return false;
}
return true;

}
</script>
<script language="JavaScript" type="text/JavaScript">
function phomod()
{
if (document.all["picmod"].style.display=="")
{

	document.all["filemod"].style.display="";
	document.all["picmod"].style.display="none";
	document.all["pho_name12"].value="";
	form1.action="z_photo_upfile1.asp";
}
else
{
	document.all["filemod"].style.display="none";
	document.all["picmod"].style.display="";
	document.all["pho_name12"].value="dis";
	form1.action="z_photo_upfile1.asp";
}
}
</script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="z_photo_upfile1.asp" method="post" enctype="multipart/form-data" name="form1" id="form1" onSubmit="return z_photo_checkform()">
<table width="97%" height="190" align="center" cellpadding="3" cellspacing="1" class=tableborder1>
	<tr> 
		<th colspan="2" height=25><div align="left">-=> 添 加 新 照 片</div></th>
	</tr>
	<%if Cint(GroupSetting(7))=0 then%>
		<tr>
			<td class=tablebody2 height=21 colspan=2 align=center>您没有在本论坛上传文件的权限</td>
		</tr>
	<%else%>
		<tr id="picmod" style="display:"> 
			<td width="10%" height=21 class=tablebody2> 照片： </td>
			<td width="90%" class=tablebody2><input type="file" name="file1" size=37></td>
		</tr>
		<tr id="filemod" style="display:none"> 
			<td width="10%" height=21 class=tablebody2> 地址： </td>
			<td width="90%" class=tablebody2><input name="pho_name12" type="text" id="pho_name12" value="dis" size="50" ></td>
		</tr>
		<tr> 
			<td height=21 class=tablebody2> 标题： </td>
			<td height=21 class=tablebody2><input name="pho_title" type="text" id="pho_title" size="50"></td>
		</tr>
		<tr> 
			<td height=21 class=tablebody2> 相册： </td>
			<td height=21 class=tablebody2> <select name="pho_groups" id="pho_groups"><% 
				dim pho_group_id,per_name
				set rs=server.createobject("adodb.recordset")
				sql="select * from z_photo_group where (pho_state = 'v' and pho_all='y' and pho_shared='y') "
				rs.open sql,z_photo_cn,1,3
				do while not rs.eof 
					pho_group_id=rs("pho_group_id")
					per_name=rs("pho_group_name")
					%><option selected value="<%= pho_group_id %>" ><%= per_name %></option><%
					rs.movenext 
				loop
				rs.close
				set rs=server.createobject("adodb.recordset")
				sql="select * from z_photo_group where (pho_state = 'v' and username='"&membername&"' and pho_all='n') "
				rs.open sql,z_photo_cn,1,3
				if not rs.eof then
					do while not rs.eof 
						pho_group_id=rs("pho_group_id")
						per_name=rs("pho_group_name")
						%><option value="<%= pho_group_id %>" ><%= per_name %></option><%
						rs.movenext 
					loop
					rs.close
					%></select> <% 
				else 
					Response.Write "</select>"
					Response.Write "你还没有建立个人相册，请到相册组管理页面中添加！"
				end if
			%></td>
    </tr>
    <tr> 
      <td height=40 valign=top class=tablebody2> 说明： </td>
      <td height=40 valign=top class=tablebody2><textarea name="pho_brief" cols="50" rows="6" wrap="VIRTUAL" id="pho_brief"></textarea></td>
    </tr>
    <tr> 
      <td valign=top class=tablebody2>&nbsp;</td>
      <td valign=top class=tablebody2><input type="hidden" name="filepath" value="photos"><input type="hidden" name="act" value="upload"><input type="submit" name="Submit" value="上 传" > <input type="reset" name="Submit2" value="取 消" >&nbsp;(<font color=<%=forum_body(8)%>>您有现金<%=mymoney%>元</font>)</td>
    </tr>
	<%end if%>
  </table>
  <%dim z_photo_totle_size,z_photo_size,z_photo_type
	set rs=server.createobject("adodb.recordset")
	sql="select * from z_photo_config where id=1"
	rs.open sql,z_photo_cn,1,3
	z_photo_totle_size=rs("z_photo_totle_size")
	z_photo_size=rs("z_photo_size")
	z_photo_type=rs("z_photo_type")
	rs.close%>
	<ul>
		<font color=<%=forum_body(8)%> align="left"> <li> 单个照片文件大小不得大于<%= z_photo_size %>字节；</li></font>
		<font color=<%=forum_body(8)%> align="left"><li> 每个用户空间限制为：<%= z_photo_totle_size %>字节；</li></font>
		<font color=<%=forum_body(8)%> align="left"> <li> 可以上传的文件类型为：<%= z_photo_type %></li></font>
		<font color=<%=forum_body(8)%> align="left"> <li> 上传文件的费用为：<%
			if isnull(myvip) or myvip<>1 then
				response.write forum_user(38)
			else
				response.write forum_user(39)
			end if
		%>元 / 1024字节</li></font>
	</ul>
</form>
</body>
</html>
<%call close_z_photo_cn()
conn.close
set conn=nothing%>
