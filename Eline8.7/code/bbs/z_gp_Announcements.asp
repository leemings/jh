<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="�������"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if founderr then
	call dvbbs_error()
else
	call AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<%select case request("action")
	case "AddAnn"		
		call AddAnn()
	case "SaveAnn"		
		call SaveAnn()
	case "EditAnn"
		call EditAnn()
	case "EditAnnInfo"
		call EditAnnInfo()		
	case "SaveEdit"		
		call SaveEdit()
	case "DelAnn"
		call DelAnn()							
	case else
		call main()
	end select%>
	</table>
<%end if
call footer()
'=====================================
sub main()
	if master then
		response.write "<tr><th colspan=2 height=25><a href=?action=AddAnn>��������</a> | <a href=?action=EditAnn>������</a></th></tr>"
	end if
	set rs=gp_conn.execute("select * from [StockNews] order by id desc")
	if rs.eof and rs.bof then
		response.Write "<tr><td class=tablebody2 colspan=2 height=""25"">��ǰû�й���</td></tr>"
	else
		do while not rs.eof%>
			<tr>
				<td class=tablebody2 colspan="2" height=25>&gt;&gt; <%=server.htmlencode(rs("title"))%> &lt;&lt;</td>
			</tr>	
			<tr>
				<td colspan="2" class=tablebody1 valign=top><br><p><blockquote><%=rs("content")%></blockquote></p></td>  
			</tr>
			<tr>
				<td class=tablebody2 align=left>&nbsp;&nbsp;&nbsp;<b>������</b>�� <%=server.htmlencode(rs("username"))%></td>
				<td class=tablebody2 align=right><b>����ʱ��</b>�� <%=rs("addtime")%>&nbsp;&nbsp;&nbsp;</td>	
			</tr>
			<%rs.movenext
		loop
	end if	
end sub
'-----------------------
sub AddAnn()
	if not master then
		Errmess=Errmess+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		call endinfo(1)
		exit sub
	end if
	response.write "<tr><th colspan=2 height=25><a href=?>�������</a> | <a href=?action=AddAnn>��������</a> | <a href=?action=EditAnn>������</a></th></tr>"%>
	<tr><td class=tablebody2 valign=middle colspan=2 align=center><B>��������</B></td></tr>
	<form action="?action=SaveAnn" method="post" name="form2"> 	
  <tr>
  	<td class=tablebody1 valign=middle><b>�û�����</b></td>
		<td class=tablebody1 valign=middle><%=membername%><INPUT name=username type=hidden value="<%=membername%>"></td>
	</tr>
  <tr>
 	 	<td class=tablebody1 valign=middle><b>�� �⣺</b><br>100������</td>
  	<td class=tablebody1 valign=middle><INPUT name="title" type=text size=60></td>
	</tr>
  <tr>
 	 	<td class=tablebody1 valign=middle><b>��ɫ��</b><br>���������ɫ</td>
		<td class=tablebody1 valign=middle>
			<SELECT name=color> 
			<option style="background-color:black;color: black" value="black" selected> black </option>
			<option style="background-color:blue;color: blue" value="blue"> blue </option>
			<option style="background-color:red;color: red" value="red"> red </option>
			<option style="background-color:olive;color: olive" value="olive"> olive </option>
			<option style="background-color:teal;color: teal" value="teal"> teal </option>
			<option style="background-color:maroon;color: maroon" value="maroon"> maroon </option>
			<option style="background-color:navy;color: navy" value="navy"> navy</option>
			<option style="background-color:gray;color: gray" value="gray"> gray</option>
			<option style="background-color:lime;color: lime" value="lime"> lime</option>
			<option style="background-color:Fuchsia;color: Fuchsia" value="Fuchsia"> Fuchsia </option>
			<option style="background-color:yellow;color: yellow" value="yellow"> yellow </option>
			<option style="background-color:green;color: green" value="green"> green</option>
			<option style="background-color:Purple;color: Purple" value="Purple"> Purple</option>
			</SELECT>
		</td>
	</tr>
  <tr>
  	<td class=tablebody1 valign=top width=30%><b>�� �ݣ�</b><br>250������<br><br><font color="<%=forum_body(8)%>">HTML�﷨ ֧��<br>UBB �﷨ ��֧��</font></td>
		<td class=tablebody1 valign=middle><textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"></textarea></td>
	</tr>
  <tr>
  	<th valign=middle colspan=2 align=center height=25><input type=submit name="submit" value="�� ��"></th>
	</tr>
	</form>
<%end sub

sub SaveAnn()
	dim username,title,content
	if not master then
		Errmess=Errmess+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		call endinfo(1)
		exit sub
	end if
	if request("username")="" then
		Errmess=Errmess+"<br>"+"<li>�����뷢���ߡ�"
		founderr=true
	else
		username=replace(request("username"),"'","")
	end if
	if request("title")="" then
		Errmess=Errmess+"<br>"+"<li>���������ű��⡣"
		founderr=true
	else
		title=replace(request("title"),"'","")
	end if
	if request("content")="" then
		Errmess=Errmess+"<br>"+"<li>�������������ݡ�"
		founderr=true
	else
		content=replace(request("content"),"'","")
	end if
	if founderr then 
		call endinfo(1)
		exit sub
	end if
	gp_conn.execute("insert into [StockNews] (title,content,[username],color,addtime) values('"&title&"','"&content&"','"&username&"','"&request.Form("color")&"',now())")	
	sucmess="<li>���Ѿ��ɹ��ķ����˹��档"
	call endinfo(1)
end sub
'---------------������--------------------
sub EditAnn()
	if not master then
		Errmess=Errmess+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		call endinfo(1)
		exit sub
	end if
	response.write "<tr><th colspan=5 height=25><a href=?>�������</a> | <a href=?action=AddAnn>��������</a> | <a href=?action=EditAnn>������</a></th></tr>"%>
  <tr>
	 	<td class=tablebody2 height=25 align=center><b>ID</b></th>
		<td class=tablebody2 align=center><b>����</b></th>
		<td class=tablebody2 align=center><b>������</b></th>
		<td class=tablebody2 align=center><b>����ʱ��</b></th>
		<td class=tablebody2 align=center><b>����</b></th>
	</tr>
	<form action="?action=DelAnn" method="post" name="form1"> 	
	<%dim hf
	set rs=gp_conn.execute("select * from [StockNews] order by id desc")
	if rs.eof and rs.bof then
		hf=0
		response.Write "<tr><td class=tablebody2 colspan=5 height=""25""><font color=gray>��ǰû�й���</font></td></tr>"
	else
		hf=1
		do while not rs.eof%>	
	    <tr align="center">
	    	<td class=tablebody1 valign=middle><%=rs("id")%></td>
				<td class=tablebody1 valign=middle align="left"><a href="?action=EditAnnInfo&id=<%=rs("id")%>"><%=server.htmlencode(rs("title"))%></a></td>
	    	<td class=tablebody1 valign=middle><%=server.htmlencode(rs("username"))%></td>
				<td class=tablebody1 valign=middle><%=rs("addtime")%></td>
	    	<td class=tablebody1 valign=middle><input type="checkbox" name="id" value="<%=rs("id")%>"></td>				
			</tr>
			<%rs.movenext
		loop
	end if
	rs.close
	if hf=1 then%>	
		<tr><th colspan=5 align=right height="25">��ѡ��Ҫɾ���Ĺ��棬<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ȫѡ <input type=submit name=Submit value=ִ��  onclick="{if(confirm('��ȷ��ִ��ɾ��������?')){this.document.form1.submit();return true;}return false;}"></th></tr>
	<%end if%>
	</form>
<%end sub
'--------------�༭����--------------
sub EditAnnInfo()
	if not master then
		Errmess=Errmess+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		call endinfo(1)
		exit sub
	end if
	if not isnumeric(request("id")) then
		Errmess=Errmess+"<br><li>����Ĳ�����"
		call endinfo(1)
		exit sub
	end if	
	response.write "<br><table border=0 cellpadding=0 cellspacing=0 align=center><tr><td><a href=?>�������</a> | <a href=?action=AddAnn>��������</a> | <a href=?action=EditAnn>������</a> | <a href=z_gp_admin_setting.asp>���й���</a></td></tr></table>"
	set rs=gp_conn.execute("select * from [StockNews] where id="&request("id"))%>
  <tr><td class=tablebody2 valign=middle colspan=2 class="admin">������й���</td></tr>
	<form action="?action=SaveEdit&id=<%=request("id")%>" method="post"> 	
  <tr>
  	<td class=tablebody1 valign=middle><b>�û�����</b></td>
		<td class=tablebody1 valign=middle><input type=text name=username value="<%=rs("username")%>"></td>
	</tr>
  <tr>
 	 	<td class=tablebody1 valign=middle><b>�� �⣺</b><br>100������</td>
  	<td class=tablebody1 valign=middle><INPUT name="title" type=text size=60 value="<%=rs("title")%>"></td>
	</tr>
	<tr>
		<td class=tablebody1 valign=middle><b>��ɫ��</b><br>���������ɫ</td>
		<td class=tablebody1 valign=middle>
			<SELECT name=color> 
			<option style="background-color:black;color: black" value="black" <%if rs("color")="black" then%> selected <%end if%>> black </option>
			<option style="background-color:blue;color: blue" value="blue" <%if rs("color")="blue" then%> selected <%end if%>> blue </option>
			<option style="background-color:red;color: red" value="red" <%if rs("color")="red" then%> selected <%end if%>> red </option>
			<option style="background-color:olive;color: olive" value="olive" <%if rs("color")="olive" then%> selected <%end if%>> olive </option>
			<option style="background-color:teal;color: teal" value="teal" <%if rs("color")="teal" then%> selected <%end if%>> teal </option>
			<option style="background-color:maroon;color: maroon" value="maroon" <%if rs("color")="maroon" then%> selected <%end if%>> maroon </option>
			<option style="background-color:navy;color: navy" value="navy" <%if rs("color")="navy" then%> selected <%end if%>> navy</option>
			<option style="background-color:gray;color: gray" value="gray" <%if rs("color")="gray" then%> selected <%end if%>> gray</option>
			<option style="background-color:lime;color: lime" value="lime" <%if rs("color")="lime" then%> selected <%end if%>> lime</option>
			<option style="background-color:Fuchsia;color: Fuchsia" value="Fuchsia" <%if rs("color")="Fuchsia" then%> selected <%end if%>> Fuchsia </option>
			<option style="background-color:yellow;color: yellow" value="yellow" <%if rs("color")="yellow" then%> selected <%end if%>> yellow </option>
			<option style="background-color:green;color: green" value="green" <%if rs("color")="green" then%> selected <%end if%>> green</option>
			<option style="background-color:Purple;color: Purple" value="Purple" <%if rs("color")="Purple" then%> selected <%end if%>> Purple</option>
			</SELECT>
		</td>
	</tr>
  <tr>
  	<td class=tablebody1 valign=top width=30%><b>�� �ݣ�</b><br>250������<br><br><font color="red">HTML�﷨ ֧��<br>UBB �﷨ ��֧��</font></td>
		<td class=tablebody1 valign=middle><textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"><%=replace(replace(rs("content"),"<br>",chr(13)),"&nbsp;"," ")%></textarea></td>
	</tr>
  <tr>
    <td class=tablebody1 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td>
	</tr>
	</form>
<%end sub

sub SaveEdit()
	dim username,title,content
	if not master then
		Errmess=Errmess+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		call endinfo(1)
		exit sub
	end if
	if not isnumeric(request("id")) then
		Errmess=Errmess+"<br><li>����Ĳ�����"
		call endinfo(1)
		exit sub
	end if	
	if request("username")="" then
		Errmess=Errmess+"<br>"+"<li>�����뷢���ߡ�"
		founderr=true
	else
		username=replace(request("username"),"'","")
	end if
	if request("title")="" then
		Errmess=Errmess+"<br>"+"<li>���������ű��⡣"
		founderr=true
	else
		title=replace(request("title"),"'","")
	end if
	if request("content")="" then
		Errmess=Errmess+"<br>"+"<li>�������������ݡ�"
		founderr=true
	else
		content=replace(request("content"),"'","")
	end if
	if founderr then 
		call endinfo(1)
		exit sub
	end if
	gp_conn.execute("update [StockNews] set title='"&title&"',content='"&content&"',[username]='"&username&"',color='"&request.Form("color")&"',addtime=now() where id="&request("id"))	
	sucmess="<li>���Ѿ��ɹ����޸��˹��档"
	rurl="?action=EditAnn"
	call endinfo(1)
end sub

sub delann()
	if not master then
		Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
		founderr=true
		exit sub
	end if
	dim delid
	delid=replace(request.form("id"),"'","")
	if delid<>"" then
		gp_conn.execute("delete from [StockNews] where id in ("&delid&")")
	end if
	sucmess="<li>ɾ������ɹ���"
	rurl="?action=EditAnn"
	call endinfo(1)
end sub
%>