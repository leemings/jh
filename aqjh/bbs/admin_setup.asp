<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")
id=int(Request("id"))


response.write "<center>"

select case Request("menu")

case "affichelist"
affichelist


case "addaffiche"
addaffiche

case "addafficheok"
sql="select * from [affiche] where id="&id&""
rs.Open sql,Conn,1,3
if rs.eof then rs.addnew
rs("title")=""&Request("title")&""
rs("content")=replace(replace(Request("content"),vbCrlf,""),"'","&#39;")
rs("username")=""&Request.Cookies("username")&""
rs("posttime")=date()
rs.update
rs.close
sql="select top 2 * from affiche order by posttime Desc"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
affiche=affiche&"<b>"&rs("title")&"</b> ("&rs("posttime")&")������"
RS.MoveNext
loop
Set Rs = Nothing
conn.execute("update [clubconfig] set affichetitle='"&affiche&"'")

%> �����ɹ�<br><br><a href=javascript:history.back()>�� ��</a><%

case "delaffiche"
conn.execute("delete from [affiche] where id="&id&"")
sql="select top 2 * from affiche order by posttime Desc"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
affiche=affiche&"<b>"&rs("title")&"</b> ("&rs("posttime")&")������"
RS.MoveNext
loop
Set Rs = Nothing
conn.execute("update [clubconfig] set affichetitle='"&affiche&"'")

%> ɾ���ɹ�<br><br><a href=javascript:history.back()>�� ��</a><%
case "variable"
variable

case "variableok"
variableok


end select

sub affichelist
%>

<a href="?menu=addaffiche">��������</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="javascript:this.location.reload()">ˢ���б�</a><br>
��<table border="0" cellpadding="5" cellspacing="1" class=a2 width=97%>
<tr>
<td width="5%" align="center" height="25" class=a1>ID</td>
<td width="35%" align="center" height="25" class=a1>����</td>
<td width="10%" align="center" height="25" class=a1>������</td>
<td width="15%" align="center" height="25" class=a1>����ʱ��</td>
<td width="15%" align="center" height="25" class=a1>����</td>
</tr>
<%
sql="select * from affiche order by posttime Desc"
rs.Open sql,Conn,1
pagesetup=20 '�趨ÿҳ����ʾ����
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '��ҳ��
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1%>
<tr>



<td height="25" align="center" bgcolor="FFFFFF"> <%=rs("id")%>
<td height="25" align="center" bgcolor="FFFFFF"><%=rs("title")%>
<td height="25" align="center" bgcolor="FFFFFF"><a target="_blank" href="Profile.asp?username=<%=rs("username")%>"><%=rs("username")%></a>
<td height="25" align="center" bgcolor="FFFFFF"><%=rs("posttime")%>
<td height="25" align="center" bgcolor="FFFFFF"><a href="?menu=addaffiche&id=<%=rs("id")%>">�޸Ĺ���</a> <a onclick=checkclick('��ȷ��Ҫɾ���ù��棿') href="?menu=delaffiche&id=<%=rs("id")%>">ɾ������</a></td>
</tr>
<%
RS.MoveNext
loop
RS.Close
%>
</table>
[<b>
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"")
</script>
</b>]

<%
end sub
sub addaffiche


if Request("id")<>empty then
sql="select * from [affiche] where id="&id&""
Set Rs=Conn.Execute(sql)
content=rs("content")
title=rs("title")
end if

%>
<form name="yuziform" method="post" action="?menu=addafficheok" onSubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=content%>'>
<input name="id" type="hidden" value='<%=id%>'>

<table cellspacing="1" cellpadding="2" width="90%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>��������</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
���⣺</td>
    <td class=a3 width="82%">
<input type="text" name="title" size="60" value="<%=title%>"></td></tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
���ݣ�</td>
    <td class=a3 width="82%" height="250">
    
    <SCRIPT src="inc/post.js"></SCRIPT>

</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<input type="submit" value=" �� �� " name="submit1">
<input name=preview type="Button" value=" Ԥ �� " onclick="Gopreview()">
<input type="reset" value=" �� �� ">
</td></tr></table></form><form name=preview action=preview.asp method=post target=preview_page><input name="content" type="hidden"></form>
<a href=javascript:history.back()>�� ��</a>
<%
end sub

sub variable
if ""&cluburl&""=empty then cluburl="http://"&Request.ServerVariables("server_name")&""&replace(Request.ServerVariables("script_name"),"admin_setup.asp","")&""
%>


<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>������Ϣ</td>
  </tr>
<form method="post" action="?menu=variableok">

<tr class=a3>
<td width="50%">�������ƣ�</td>
<td><input size="30" name="clubname" value="<%=clubname%>"></td>
</tr>
<tr class=a4>
<td>����URL��<br>
<script>
autourl="http://<%=Request.ServerVariables("server_name")%><%=replace(Request.ServerVariables("script_name"),"admin_setup.asp","")%>"

if (autourl != '<%=cluburl%>'){
document.write("ϵͳ�Զ���⵽��URL��<font color=FF0000>"+autourl+"</font>");
}
</script>

</td>
<td><input size="30" name="cluburl" value="<%=cluburl%>"></td>
</tr>
<tr class=a3>
<td>��ҳ���ƣ�</td>
<td><input size="30" name="homename" value="<%=homename%>"></td>
</tr>
<tr class=a4>
<td>��ҳ��ַ��</td>
<td><input size="30" value="<%=homeurl%>" name="homeurl"></td>
</tr>

</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25 class=a1>
    <td class=a4 align=middle colspan=2>��Ϣ����</td>
  </tr>

  
  <tr class=a3>
<td width="50%">��ҳ��ʾ��̳��ȣ�</td>
<td>
<input size="4" value="<%=floor%>" name="floor"> ����[Ĭ��:2]</tr>
	</tr>



    <tr class=a3>
<td>�û����������</td>
<td><input size="4" value="<%=PostTime%>" name="PostTime"> �롡[Ĭ��:30]</td>
</tr>
  <tr class=a4>
<td>�ű���ʱʱ�䣺</td>
<td><input size="4" value="<%=Timeout%>" name="Timeout"> �롡[Ĭ��:60]</td>
</tr>

  
  <tr class=a3>
<td width="50%">ͳ���û�����ʱ�䣺</td>
<td><input size="4" value="<%=OnlineTime%>" name="OnlineTime"> �롡[Ĭ��:1200]</td>
	</tr>
 
  
  <tr class=a4>
<td width="50%">ע���೤ʱ����ܷ�������</td>
<td><input size="4" value="<%=Reg10%>" name="Reg10"> ���ӡ�[Ĭ��:10]</td>
	</tr>
 
  <tr class=a3>
<td>������̳�������ƣ������һ��վ���ж����̳����ĳɲ�ͬ���ƣ�</td>
<td>
<input size="8" value="<%=CacheName%>" name="CacheName">��[Ĭ��:bbsxp]</td>
</tr>
  
<tr class=a4>
<td>Ĭ�Ϸ�����ã���ָ��<font color="#FF0000">images/skins/</font>Ŀ¼�µ�Ŀ¼�����ɣ�</td>
<td>
<input size="8" value="<%=style%>" name="style">��[Ĭ��:1]</td>
</tr>



</table>




<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�ʼ���Ϣ</td>
  </tr>
  

<tr class=a3>
<td width="50%"0>�����ʼ������</td>
<td>
<select name="selectmail">
<option value="">�ر�</option>
<option value="JMail" <%if selectmail="JMail" then%>selected<%end if%>>JMail.Message</option>
<option value="CDONTS" <%if selectmail="CDONTS" then%>selected<%end if%>>CDONTS.NewMail</option>
</select>
</td>
</tr>
<tr class=a4>
<td>SMTP Server��ַ��</td>
<td><input size="30" value="<%=smtp%>" name="smtp"></td>
</tr>
<tr class=a3>
<td>�ʼ���������¼����</td>
<td><input size="30" value="<%=MailServerusername%>" name="MailServerusername"></td>
</tr>
<tr class=a4>
<td>�ʼ���������¼���룺</td>
<td>
<input size="30" value="<%=MailServerPassword%>" name="MailServerPassword" type="password"></td>
</tr>
<tr class=a3>
<td>������Email��ַ��</td>
<td><input size="30" value="<%=smtpmail%>" name="smtpmail"></td>
</tr>  
</table>




<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�û�ѡ��</td>
  </tr>


<tr class=a3>
<td width="50%">���û�ע�᣺</td>
<td>
<input type=radio name="CloseRegUser" value="1" <%if CloseRegUser=1 then%>checked<%end if%> id=CloseRegUser><label for=CloseRegUser>�ر�</label>
<input type=radio name="CloseRegUser" value="0" <%if CloseRegUser=0 then%>checked<%end if%> id=CloseRegUser_1><label for=CloseRegUser_1>����</label>
</td>
</tr>



<tr class=a4>
<td>һ��Emailֻ��ע��һ���ʺ�</td>
<td>
<input type=radio name="RegOnlyMail" value="0" <%if RegOnlyMail=0 then%>checked<%end if%> id=RegOnlyMail><label for=RegOnlyMail>�ر�</label>
<input type=radio name="RegOnlyMail" value="1" <%if RegOnlyMail=1 then%>checked<%end if%> id=RegOnlyMail_1><label for=RegOnlyMail_1>��</label>
</td>
</tr>

<tr class=a3>
<td width="50%">ע���û�����ͨ��Email���ͣ�<br><font color="FF0000">����������֧���ʼ����͹���</font></td>
<td>
<input type=radio name="SendPassword" value="0" <%if SendPassword=0 then%>checked<%end if%> id=SendPassword><label for=SendPassword>�ر�</label>
<input type=radio name="SendPassword" value="1" <%if SendPassword=1 then%>checked<%end if%> id=SendPassword_1><label for=SendPassword_1>��</label>
</td>
</tr>




</table>


  <br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>����ѡ��</td>
  </tr>
<tr class=a4>
<td width="50%">�����������Ľ����</td>
<td>
&nbsp;<input size="4" value="<%=MaxSearch%>" name="MaxSearch">&nbsp; [Ĭ��:500]</td>
</tr>
<tr class=a4>
<td width="50%">�Ƿ���������������</td>
<td>
<input type=radio name="searchcontent" value="0" <%if searchcontent=0 then%>checked<%end if%> id=searchcontent><label for=searchcontent>�ر�</label>
<input type=radio name="searchcontent" value="1" <%if searchcontent=1 then%>checked<%end if%> id=searchcontent_1><label for=searchcontent_1>��</label>
</td>
</tr>
  
</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>��̳ѡ��</td>
  </tr>
  

</tr>
  

<tr class=a3>
<td width="50%">�����Ƿ���Ҫ�������������ʾ��</td>
<td>
<input type=radio name="ActivationForum" value="0" <%if ActivationForum=0 then%>checked<%end if%> id=ActivationForum><label for=ActivationForum>��</label>&nbsp;&nbsp;
<input type=radio name="ActivationForum" value="1" <%if ActivationForum=1 then%>checked<%end if%> id=ActivationForum_1><label for=ActivationForum_1>��</label>
</td>
</tr>

<tr class=a4>
<td width="50%">�û��Ƿ���Ҫ����������ܵ�¼��</td>
<td>
<input type=radio name="ActivationUser" value="1" <%if ActivationUser=1 then%>checked<%end if%> id=ActivationUser><label for=ActivationUser>��</label>&nbsp;&nbsp;
<input type=radio name="ActivationUser" value="0" <%if ActivationUser=0 then%>checked<%end if%> id=ActivationUser_1><label for=ActivationUser_1>��</label>
</td>
</tr>





<tr class=a3>
<td width="50%">����������̳���ܣ�</td>
<td>
<input type=radio name="apply" value="0" <%if apply=0 then%>checked<%end if%> id=apply><label for=apply>�ر�</label>
<input type=radio name="apply" value="1" <%if apply=1 then%>checked<%end if%> id=apply_1><label for=apply_1>����</label>
</td>
</tr>


  <tr class=a4>
<td width="50%">�����̳��������̳ӵ��ͬ�����ܣ�
���磺<font color="FF0000">���������</font>��</td>
<td>
<input type=radio name="SortShowForum" value="0" <%if SortShowForum=0 then%>checked<%end if%> id=SortShowForum><label for=SortShowForum>�ر�</label>
<input type=radio name="SortShowForum" value="1" <%if SortShowForum=1 then%>checked<%end if%> id=SortShowForum_1><label for=SortShowForum_1>����</label>
</td></tr>


<tr class=a3>
<td width="50%">����Դ��⣨��ֹ���ⲿ�ύ���ݣ���</td>
<td>
<input type=radio name="StopOutPost" value="0" <%if StopOutPost=0 then%>checked<%end if%> id=StopOutPost><label for=StopOutPost>�ر�</label>
<input type=radio name="StopOutPost" value="1" <%if StopOutPost=1 then%>checked<%end if%> id=StopOutPost_1><label for=StopOutPost_1>��</label>
</td>
</tr>



<tr class=a4>
<td>��ҳ��ʾ��̳�б���ʽ��</td>
<td>
<input type=radio name="ListStyle" value="0" <%if ListStyle=0 then%>checked<%end if%> id=ListStyle><label for=ListStyle>���</label>
<input type=radio name="ListStyle" value="1" <%if ListStyle=1 then%>checked<%end if%> id=ListStyle_1><label for=ListStyle_1>��ϸ</label>
</td>
</tr>




</table>

<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�ϴ�ѡ��</td>
  </tr>
  
<tr class=a4>
<td width="50%">ѡ���ϴ������</td>
<td>
<select name="selectup">
<option value="">�ر�</option>
<option value="FSO" <%if selectup="FSO" then%>selected<%end if%>>FSO</option>
<option value="SA-FileUp" <%if selectup="SA-FileUp" then%>selected<%end if%>>SA-FileUp</option>
</select></td>
</tr>
  
<tr class=a3>
<td width="50%">����ͷ���ļ��Ĵ�С��</td>
<td>
<input size="6" value="<%=MaxFace%>" name="MaxFace" readonly> byte&nbsp; [Ĭ��:10240]</td>
</tr>

<tr class=a4>
<td>������Ƭ�ļ��Ĵ�С��</td>
<td>
<input size="6" value="<%=MaxPhoto%>" name="MaxPhoto" readonly> byte&nbsp; [Ĭ��:30720]</td>
</tr>
  
 <tr class=a3>
<td>�������ļ��Ĵ�С��</td>
<td>
<input size="6" value="<%=MaxFile%>" name="MaxFile" readonly> byte&nbsp; [Ĭ��:102400]</td>
</tr>
   
  
   <tr class=a4>
<td>�������ļ������ͣ�<br>���磺<font color="FF0000">gif|jpg|png|bmp|swf|txt|mid|doc|xls|zip|rar</font></td>
<td>
<input size="40" value="<%=UpFileGenre%>" name="UpFileGenre" readonly> </td>
</tr>

</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25 class=a1>
    <td class=a4 align=middle colspan=2>��������</td>
  </tr>

  <tr class=a4>
<td width="50%">���������֣��������á�|���ָ�����<br>���磺<font color="FF0000">fuck|shit|����</font></td>
<td><input size="40" value="<%=badwords%>" name="badwords"></td>
</tr>

  <tr class=a3>
<td width="50%">�����û����ӣ����û����á�|���ָ�����<br>���磺<font color="FF0000">yuzi|ԣԣ</font></td>
<td><input size="40" value="<%=badlist%>" name="badlist"></td>
</tr>

<tr class=a4>
<td width="50%">��ֹIP��ַ������̳����IP���á�|���ָ�����<br>���磺<font color="FF0000">127.0.0.1|192.168.0.1</font></td>
<td><input size="40" value="<%=badip%>" name="badip"></td>
</tr>

</table>


<br>






<input type="submit" value=" �� �� "></form>
<%
end sub


sub variableok

if Request("clubname")="" then error2("��������������")
if Request("style")="" then error2("Ĭ�Ϸ����û������")
if Request("SendPassword")="1" and Request("selectmail")="" then error2("��ѡ����ע���û�����ͨ��Email���ͣ�������û���趨�����ʼ����")


filtrate=split(Request("badip"),"|")
for i = 0 to ubound(filtrate)
if instr("|"&Request.ServerVariables("REMOTE_ADDR")&"","|"&filtrate(i)&"") > 0 then error2("�����������IP��ַ�Ƿ���ȷ")
next




rs.Open "clubconfig",Conn,1,3
rs("badwords")=Request("badwords")
rs("clubname")=Request("clubname")
rs("cluburl")=Request("cluburl")
rs("homename")=Request("homename")
rs("homeurl")=Request("homeurl")
rs("selectmail")=Request("selectmail")
rs("smtp")=Request("smtp")
rs("smtpmail")=Request("smtpmail")
rs("MailServerusername")=Request("MailServerusername")
rs("MailServerPassword")=Request("MailServerPassword")
rs("badlist")=Request("badlist")
rs("badip")=Request("badip")
rs("style")=Request("style")
rs("CacheName")=Request("CacheName")
rs("selectup")=Request("selectup")
rs("UpFileGenre")=Request("UpFileGenre")
rs("allclass")=""&int(Request("floor"))&"\"&int(Request("CloseRegUser"))&"\"&int(Request("Reg10"))&"\"&int(Request("RegOnlyMail"))&"\"&int(Request("SendPassword"))&"\"&int(Request("apply"))&"\"&int(Request("Timeout"))&"\"&int(Request("OnlineTime"))&"\"&int(Request("MaxFace"))&"\"&int(Request("MaxPhoto"))&"\"&int(Request("MaxFile"))&"\"&int(Request("searchcontent"))&"\"&int(Request("MaxSearch"))&"\"&int(Request("ActivationForum"))&"\"&int(Request("ActivationUser"))&"\"&int(Request("SortShowForum"))&"\"&int(Request("PostTime"))&"\"&int(Request("ListStyle"))&"\"&int(Request("StopOutPost"))&""
rs.update

rs.close



on error resume next
if Request("selectmail")="JMail" then
Set JMail=Server.CreateObject("JMail.Message")
If -2147221005 = Err Then error2("����������֧�� JMail.Message �������رշ����ʼ����ܣ�")
elseif Request("selectmail")="CDONTS" then
Set MailObject = Server.CreateObject("CDONTS.NewMail")
If -2147221005 = Err Then error2("����������֧�� CDONTS.NewMail �������رշ����ʼ����ܣ�")
end if


%>
���³ɹ�<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub



htmlend

%>