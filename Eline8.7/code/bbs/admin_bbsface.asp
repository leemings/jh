<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
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
	actname="ע��ͷ��"
	picfilename="image"
case 2
	bbspicmun=Forum_PostFaceNum
	bbspicurl=Forum_info(8)
	connfile=Forum_PostFace
	actname="��������ͼƬ"
	picfilename="face"
case 3
	bbspicmun=Forum_emotNum
	bbspicurl=Forum_info(10)
	connfile=Forum_emot
	actname="��������ͼƬ"
	picfilename="em"
case else
	bbspicmun=Forum_userfaceNum
	bbspicurl=Forum_info(11)
	connfile=Forum_userface
	actname="ע��ͷ��"
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
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	if request("Submit")="��������" then
		call saveconst()
	elseif request("Submit")="�ָ�Ĭ������" then
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
		<td height="23" colspan="4" ><B>˵��</B>��<br>�١�����ͼƬ����������̳<%=bbspicurl%>Ŀ¼�У���Ҫ����Ҳ�뽫ͼƬ���ڸ�Ŀ¼<br>�ڡ��ұ߸�ѡ��Ϊɾ��ѡ����ѡ���㱣�����ã���ɾ����ӦͼƬ<BR>�ۡ�������޸��ļ����������޸���Ӧѡ���ֱ�ӵ���������ö�����ѡ���ұ߸�ѡ��</td>
	</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr bgcolor=<%=Forum_body(3)%>> 
		<th height="23" colspan="4" align=left><%=actname%>�������� ��Ŀǰ����<%=count%>��<%=actname%>ͼƬ���ļ��У�<%=bbspicurl%>��</th>
	</tr>
	<tr bgcolor=<%=Forum_body(3)%>> 
		<td width="20%"  align=left class=forumrow>���ӵ��ļ�����</td>
		<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWFILENAME" value="<%=newfilename%>">��<font color=red>�������Ĭ�ϣ����Ӻ����Ӧ���ļ����ϴ�����Ŀ¼�¡�</font>��</td>
	</tr>
	<tr bgcolor=<%=Forum_body(3)%>> 
		<td width="20%"  align=left class=forumrow>����������Ŀ��</td>
		<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWNUM" value="<%=newnum%>"> <input type="submit" name="Submit" value="����"></td>
	</tr>
	<%IF REQUEST("Submit")="����" and REQUEST("Newnum")<>"" and request("Newnum")<>0 then
		newnum=REQUEST("Newnum")
		for i=count+1 to count+newnum%>
			<tr>
				<td width="20%" class=forumRowHighlight><%=actname%>ID��<input type=hidden name="face_id<%=i%>" size="10" value="<%=i%>"><%=i%></td>
				<td width="75%" class=forumRowHighlight colspan="2">�����ӵ��ļ���<input  type="text" name="userface<%=i%>" value="<%=newfilename&i+1%>.gif"></td>
				<td width="5%" class=forumRowHighlight> <input type="checkbox" name="delid<%=i%>" value="<%=i+1%>"></td>
			</tr>
		<% next 
	end if%>
	<tr>
		<th width="20%" class=forumrow>�ļ�</th>
		<th width="45%" class=forumrow>�ļ���</th>
		<th width="30%" class=forumrow>ͼƬ</th>
		<th width="5%" class=forumrow>ɾ��</th>
	</tr>
	<% for i=0 to bbspicmun%>
		<tr>
			<td width="20%" class=forumrow>�ļ�����<input type=hidden  name="face_id<%=i%>" size="10" value="<%=i%>"></td>
			<td width="45%" class=forumrow>&nbsp;<input  type="text" name="userface<%=i%>" value="<%=connfile(i)%>"></td>
			<td width="30%" class=forumrow>&nbsp;&nbsp;<img src=<%=bbspicurl%><%=connfile(i)%>></td>
			<td width="5%" class=forumrow> <input type="checkbox" name="delid<%=i%>" value="<%=i+1%>"></td>
		</tr>
	<% next %>
	<tr> 
		<td  colspan="4" class=forumrow> <B>ע��</B>���ұ߸�ѡ��Ϊɾ��ѡ����ѡ���㱣�����ã���ɾ����ӦͼƬ<BR>������޸��ļ����������޸���Ӧѡ���ֱ�ӵ���������ö�����ѡ���ұ߸�ѡ��</td>
	</tr>
	<tr> 
		<td  colspan="4" class=forumrow align="center">  ɾ��ѡ�ɾ����ѡ��ʵ���ļ���<font color=red>��ҪFSO֧�ֹ���</font>������<input type=radio name=setfso value=1 >��<input type=radio name=setfso value=0 > ��ѡ��Ҫɾ�����ļ���<input type="checkbox" name=chkall value=on onclick="CheckAll(this.form)">ȫѡ <BR><input type="submit" name="Submit" value="��������">   <input type="submit" name="Submit" value="�ָ�Ĭ������"></td>
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
		Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
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
						response.write "ɾ��"&filepaths
					else
						response.write "δ�ҵ�"&filepaths
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
		response.write "<br>���óɹ���<a href="&Request.ServerVariables("HTTP_REFERER")&" >�������</a>"
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
	response.write "<br>���óɹ���<a href="&Request.ServerVariables("HTTP_REFERER")&" >�������</a>"
end sub

%>

