<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
dim face_id,count
dim newnum,newfilename
dim orders,bbspicmun,bbspicurl,picfilename,actname,connfile,upconfig

if not isInteger(request("orders")) or request("orders")="" then
	orders=1
else
	orders=request("orders")
end if
select case orders
case 1     
	bbspicmun=Forum_userfaceNum
	bbspicurl=Forum_info(11)
	connfile=Forum_userface
	actname="注册头像"
	picfilename="image"
case 2
	bbspicmun=Forum_PostFaceNum
	bbspicurl=Forum_info(8)
	connfile=Forum_PostFace
	actname="发帖心情图片"
	picfilename="face"
case 3
	bbspicmun=Forum_emotNum
	bbspicurl=Forum_info(10)
	connfile=Forum_emot
	actname="发帖表情图片"
	picfilename="em"
case else
	bbspicmun=Forum_userfaceNum
	bbspicurl=Forum_info(11)
	connfile=Forum_userface
	actname="注册头像"
	picfilename="image"
end select

if trim(request("newfilename"))<>"" then
	newfilename=trim(request("newfilename"))
else
	newfilename=picfilename
end if 

if bbspicmun<0 then 
	count=0
else
	count=bbspicmun
end if

if REQUEST("Newnum")<>"" and request("Newnum")<>0 then
	newnum=REQUEST("Newnum")
else
	newnum=0
end if
dim admin_flag
admin_flag="73,74,75"
if not master or instr(session("flag"),"73")=0 or instr(session("flag"),"74")=0 or instr(session("flag"),"75")=0 then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	if request("Submit")="保存设置" then
		call saveconst()
	elseif request("Submit")="恢复默认设置" then
		call savedefault()
	else
		call consted()
	end if
	if founderr then call dvbbs_error()
	conn.close
	set conn=nothing
end if

sub consted()
	dim sel%>
<form method="POST" action=?orders=<%=request("orders")%>&skinid=<%=skinid%> name="bbspic" >
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
	<tr> 
		<td height="23" colspan="4" ><B>说明</B>：<br>①、以下图片均保存于论坛<%=bbspicurl%>目录中，如要更换也请将图片放于该目录<br>②、右边复选框为删除选项，如果选择后点保存设置，则删除相应图片<BR>③、如仅仅修改文件名，可在修改相应选项后直接点击保存设置而不用选择右边复选框</td>
	</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr bgcolor=<%=Forum_body(3)%>> 
		<th height="23" colspan="4" align=left><%=actname%>管理设置 （目前共有<%=count%>个<%=actname%>图片在文件夹：<%=bbspicurl%>）</th>
	</tr>
	<tr bgcolor=<%=Forum_body(3)%>> 
		<td width="20%"  align=left class=forumrow>增加的文件名：</td>
		<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWFILENAME" value="<%=newfilename%>">（<font color=red>建议采用默认，增加后把相应的文件名上传到该目录下。</font>）</td>
	</tr>
	<tr bgcolor=<%=Forum_body(3)%>> 
		<td width="20%"  align=left class=forumrow>批量增加数目：</td>
		<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWNUM" value="<%=newnum%>"> <input type="submit" name="Submit" value="增加"></td>
	</tr>
	<%IF REQUEST("Submit")="增加" and REQUEST("Newnum")<>"" and request("Newnum")<>0 then
		newnum=REQUEST("Newnum")
		for i=count+1 to count+newnum%>
			<tr>
				<td width="20%" class=forumRowHighlight><%=actname%>ID：<input type=hidden name="face_id<%=i%>" size="10" value="<%=i%>"><%=i%></td>
				<td width="75%" class=forumRowHighlight colspan="2">新增加的文件：<input  type="text" name="userface<%=i%>" value="<%=newfilename&i+1%>.gif"></td>
				<td width="5%" class=forumRowHighlight> <input type="checkbox" name="delid<%=i%>" value="<%=i+1%>"></td>
			</tr>
		<% next 
	end if%>
	<tr>
		<th width="20%" class=forumrow>文件</th>
		<th width="45%" class=forumrow>文件名</th>
		<th width="30%" class=forumrow>图片</th>
		<th width="5%" class=forumrow>删除</th>
	</tr>
	<% for i=0 to bbspicmun%>
		<tr>
			<td width="20%" class=forumrow>文件名：<input type=hidden  name="face_id<%=i%>" size="10" value="<%=i%>"></td>
			<td width="45%" class=forumrow>&nbsp;<input  type="text" name="userface<%=i%>" value="<%=connfile(i)%>"></td>
			<td width="30%" class=forumrow>&nbsp;&nbsp;<img src=<%=bbspicurl%><%=connfile(i)%>></td>
			<td width="5%" class=forumrow> <input type="checkbox" name="delid<%=i%>" value="<%=i+1%>"></td>
		</tr>
	<% next %>
	<tr> 
		<td  colspan="4" class=forumrow> <B>注意</B>：右边复选框为删除选项，如果选择后点保存设置，则删除相应图片<BR>如仅仅修改文件名，可在修改相应选项后直接点击保存设置而不用选择右边复选框</td>
	</tr>
	<tr> 
		<td  colspan="4" class=forumrow align="center">  删除选项：删除所选的实际文件（<font color=red>需要FSO支持功能</font>）：是<input type=radio name=setfso value=1 >否<input type=radio name=setfso value=0 > 请选择要删除的文件，<input type="checkbox" name=chkall value=on onclick="CheckAll(this.form)">全选 <BR><input type="submit" name="Submit" value="保存设置">   <input type="submit" name="Submit" value="恢复默认设置"></td>
	</tr>
</table>
<BR><BR>
</form>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
<%
end sub

sub saveconst()
	dim f_userface,formname,d_elid,faceid
	dim filepaths,objFSO,upface
	
	if trim(request("skinid"))="" and trim(request("boardid"))="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
	else
		for i=0 to count+newnum
			faceid="face_id"&i
			d_elid="delid"&i
			formname="userface"&i
			if cint(request.Form(d_elid))=0 then 
				if f_userface<>"" then f_userface=f_userface&"|"
				f_userface=f_userface&request.Form(formname)
			else
				upface=bbspicurl&request.Form(formname)
				if  request("setfso")=1 then
					filepaths=Server.MapPath(""&upface&"")
					Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
					if objFSO.fileExists(filepaths) then
						objFSO.DeleteFile(filepaths)
						response.write "删除"&filepaths
					else
						response.write "未找到"&filepaths
					end if
				end if
			end if
		next
		if request("skinid")<>"" then
			select case orders
			case 1     
				upconfig=" Forum_userface='"&f_userface& "' "
			case 2
				upconfig=" Forum_PostFace='"&f_userface& "' "
			case 3
				upconfig=" Forum_emot='"&f_userface& "' "
			case else
				upconfig=" Forum_userface='"&f_userface& "' "
			end select
			sql = "update config set "&upconfig&" "
			conn.execute(sql)
		end if
		response.write "<br>设置成功。<a href="&Request.ServerVariables("HTTP_REFERER")&" >点击返回</a>"
	end if
end sub

sub savedefault()
	dim userface,PostFace,emot
	select case orders
	case 1     
		for i=1 to 60
		userface=userface&"face"&i&".gif|"
		next
		sql = "update config set Forum_userface='"&userface&"' "
		conn.execute(sql)
	case 2
		for i=1 to 18
		PostFace=PostFace&"face"&i&".gif|"
		next
		sql = "update config set Forum_PostFace='"&PostFace&"' "
		conn.execute(sql)
	case 3        
		for i=1 to 9
		emot=emot&"em0"&i&".gif|"
		next 
		for i=10 to 28
		emot=emot&"em"&i&".gif|"
		next
		sql = "update config set Forum_emot='"&emot&"'  "
		conn.execute(sql)
	case else
		for i=1 to 60
		userface=userface&"face"&i&".gif|"
		next
		sql = "update config set Forum_userface='"&userface&"' "
		conn.execute(sql)
	end select
	response.write "<br>设置成功。<a href="&Request.ServerVariables("HTTP_REFERER")&" >点击返回</a>"
end sub

%>

