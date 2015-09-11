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
affiche=affiche&"<b>"&rs("title")&"</b> ("&rs("posttime")&")　　　"
RS.MoveNext
loop
Set Rs = Nothing
conn.execute("update [clubconfig] set affichetitle='"&affiche&"'")

%> 发布成功<br><br><a href=javascript:history.back()>返 回</a><%

case "delaffiche"
conn.execute("delete from [affiche] where id="&id&"")
sql="select top 2 * from affiche order by posttime Desc"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
affiche=affiche&"<b>"&rs("title")&"</b> ("&rs("posttime")&")　　　"
RS.MoveNext
loop
Set Rs = Nothing
conn.execute("update [clubconfig] set affichetitle='"&affiche&"'")

%> 删除成功<br><br><a href=javascript:history.back()>返 回</a><%
case "variable"
variable

case "variableok"
variableok


end select

sub affichelist
%>

<a href="?menu=addaffiche">发布公告</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="javascript:this.location.reload()">刷新列表</a><br>
　<table border="0" cellpadding="5" cellspacing="1" class=a2 width=97%>
<tr>
<td width="5%" align="center" height="25" class=a1>ID</td>
<td width="35%" align="center" height="25" class=a1>标题</td>
<td width="10%" align="center" height="25" class=a1>发布人</td>
<td width="15%" align="center" height="25" class=a1>发布时间</td>
<td width="15%" align="center" height="25" class=a1>管理</td>
</tr>
<%
sql="select * from affiche order by posttime Desc"
rs.Open sql,Conn,1
pagesetup=20 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1%>
<tr>



<td height="25" align="center" bgcolor="FFFFFF"> <%=rs("id")%>
<td height="25" align="center" bgcolor="FFFFFF"><%=rs("title")%>
<td height="25" align="center" bgcolor="FFFFFF"><a target="_blank" href="Profile.asp?username=<%=rs("username")%>"><%=rs("username")%></a>
<td height="25" align="center" bgcolor="FFFFFF"><%=rs("posttime")%>
<td height="25" align="center" bgcolor="FFFFFF"><a href="?menu=addaffiche&id=<%=rs("id")%>">修改公告</a> <a onclick=checkclick('您确定要删除该公告？') href="?menu=delaffiche&id=<%=rs("id")%>">删除公告</a></td>
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
    <td class=a1 align=middle colspan=2>发布公告</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
标题：</td>
    <td class=a3 width="82%">
<input type="text" name="title" size="60" value="<%=title%>"></td></tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
内容：</td>
    <td class=a3 width="82%" height="250">
    
    <SCRIPT src="inc/post.js"></SCRIPT>

</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<input type="submit" value=" 发 送 " name="submit1">
<input name=preview type="Button" value=" 预 览 " onclick="Gopreview()">
<input type="reset" value=" 重 置 ">
</td></tr></table></form><form name=preview action=preview.asp method=post target=preview_page><input name="content" type="hidden"></form>
<a href=javascript:history.back()>返 回</a>
<%
end sub

sub variable
if ""&cluburl&""=empty then cluburl="http://"&Request.ServerVariables("server_name")&""&replace(Request.ServerVariables("script_name"),"admin_setup.asp","")&""
%>


<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>基本信息</td>
  </tr>
<form method="post" action="?menu=variableok">

<tr class=a3>
<td width="50%">社区名称：</td>
<td><input size="30" name="clubname" value="<%=clubname%>"></td>
</tr>
<tr class=a4>
<td>社区URL：<br>
<script>
autourl="http://<%=Request.ServerVariables("server_name")%><%=replace(Request.ServerVariables("script_name"),"admin_setup.asp","")%>"

if (autourl != '<%=cluburl%>'){
document.write("系统自动检测到的URL：<font color=FF0000>"+autourl+"</font>");
}
</script>

</td>
<td><input size="30" name="cluburl" value="<%=cluburl%>"></td>
</tr>
<tr class=a3>
<td>主页名称：</td>
<td><input size="30" name="homename" value="<%=homename%>"></td>
</tr>
<tr class=a4>
<td>主页地址：</td>
<td><input size="30" value="<%=homeurl%>" name="homeurl"></td>
</tr>

</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25 class=a1>
    <td class=a4 align=middle colspan=2>信息设置</td>
  </tr>

  
  <tr class=a3>
<td width="50%">首页显示论坛深度：</td>
<td>
<input size="4" value="<%=floor%>" name="floor"> 级　[默认:2]</tr>
	</tr>



    <tr class=a3>
<td>用户发帖间隔：</td>
<td><input size="4" value="<%=PostTime%>" name="PostTime"> 秒　[默认:30]</td>
</tr>
  <tr class=a4>
<td>脚本超时时间：</td>
<td><input size="4" value="<%=Timeout%>" name="Timeout"> 秒　[默认:60]</td>
</tr>

  
  <tr class=a3>
<td width="50%">统计用户在线时间：</td>
<td><input size="4" value="<%=OnlineTime%>" name="OnlineTime"> 秒　[默认:1200]</td>
	</tr>
 
  
  <tr class=a4>
<td width="50%">注册后多长时间才能发表帖子</td>
<td><input size="4" value="<%=Reg10%>" name="Reg10"> 分钟　[默认:10]</td>
	</tr>
 
  <tr class=a3>
<td>定义论坛缓存名称：（如果一个站点有多个论坛请更改成不同名称）</td>
<td>
<input size="8" value="<%=CacheName%>" name="CacheName">　[默认:bbsxp]</td>
</tr>
  
<tr class=a4>
<td>默认风格设置：（指定<font color="#FF0000">images/skins/</font>目录下的目录名即可）</td>
<td>
<input size="8" value="<%=style%>" name="style">　[默认:1]</td>
</tr>



</table>




<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>邮件信息</td>
  </tr>
  

<tr class=a3>
<td width="50%"0>发送邮件组件：</td>
<td>
<select name="selectmail">
<option value="">关闭</option>
<option value="JMail" <%if selectmail="JMail" then%>selected<%end if%>>JMail.Message</option>
<option value="CDONTS" <%if selectmail="CDONTS" then%>selected<%end if%>>CDONTS.NewMail</option>
</select>
</td>
</tr>
<tr class=a4>
<td>SMTP Server地址：</td>
<td><input size="30" value="<%=smtp%>" name="smtp"></td>
</tr>
<tr class=a3>
<td>邮件服务器登录名：</td>
<td><input size="30" value="<%=MailServerusername%>" name="MailServerusername"></td>
</tr>
<tr class=a4>
<td>邮件服务器登录密码：</td>
<td>
<input size="30" value="<%=MailServerPassword%>" name="MailServerPassword" type="password"></td>
</tr>
<tr class=a3>
<td>发件人Email地址：</td>
<td><input size="30" value="<%=smtpmail%>" name="smtpmail"></td>
</tr>  
</table>




<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>用户选项</td>
  </tr>


<tr class=a3>
<td width="50%">新用户注册：</td>
<td>
<input type=radio name="CloseRegUser" value="1" <%if CloseRegUser=1 then%>checked<%end if%> id=CloseRegUser><label for=CloseRegUser>关闭</label>
<input type=radio name="CloseRegUser" value="0" <%if CloseRegUser=0 then%>checked<%end if%> id=CloseRegUser_1><label for=CloseRegUser_1>开放</label>
</td>
</tr>



<tr class=a4>
<td>一个Email只能注册一个帐号</td>
<td>
<input type=radio name="RegOnlyMail" value="0" <%if RegOnlyMail=0 then%>checked<%end if%> id=RegOnlyMail><label for=RegOnlyMail>关闭</label>
<input type=radio name="RegOnlyMail" value="1" <%if RegOnlyMail=1 then%>checked<%end if%> id=RegOnlyMail_1><label for=RegOnlyMail_1>打开</label>
</td>
</tr>

<tr class=a3>
<td width="50%">注册用户密码通过Email发送：<br><font color="FF0000">服务器必须支持邮件发送功能</font></td>
<td>
<input type=radio name="SendPassword" value="0" <%if SendPassword=0 then%>checked<%end if%> id=SendPassword><label for=SendPassword>关闭</label>
<input type=radio name="SendPassword" value="1" <%if SendPassword=1 then%>checked<%end if%> id=SendPassword_1><label for=SendPassword_1>打开</label>
</td>
</tr>




</table>


  <br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>搜索选项</td>
  </tr>
<tr class=a4>
<td width="50%">搜索返回最多的结果数</td>
<td>
&nbsp;<input size="4" value="<%=MaxSearch%>" name="MaxSearch">&nbsp; [默认:500]</td>
</tr>
<tr class=a4>
<td width="50%">是否开启帖子内容搜索</td>
<td>
<input type=radio name="searchcontent" value="0" <%if searchcontent=0 then%>checked<%end if%> id=searchcontent><label for=searchcontent>关闭</label>
<input type=radio name="searchcontent" value="1" <%if searchcontent=1 then%>checked<%end if%> id=searchcontent_1><label for=searchcontent_1>打开</label>
</td>
</tr>
  
</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>论坛选项</td>
  </tr>
  

</tr>
  

<tr class=a3>
<td width="50%">帖子是否需要经过激活才能显示：</td>
<td>
<input type=radio name="ActivationForum" value="0" <%if ActivationForum=0 then%>checked<%end if%> id=ActivationForum><label for=ActivationForum>否</label>&nbsp;&nbsp;
<input type=radio name="ActivationForum" value="1" <%if ActivationForum=1 then%>checked<%end if%> id=ActivationForum_1><label for=ActivationForum_1>是</label>
</td>
</tr>

<tr class=a4>
<td width="50%">用户是否需要经过激活才能登录：</td>
<td>
<input type=radio name="ActivationUser" value="1" <%if ActivationUser=1 then%>checked<%end if%> id=ActivationUser><label for=ActivationUser>否</label>&nbsp;&nbsp;
<input type=radio name="ActivationUser" value="0" <%if ActivationUser=0 then%>checked<%end if%> id=ActivationUser_1><label for=ActivationUser_1>是</label>
</td>
</tr>





<tr class=a3>
<td width="50%">开放申请论坛功能：</td>
<td>
<input type=radio name="apply" value="0" <%if apply=0 then%>checked<%end if%> id=apply><label for=apply>关闭</label>
<input type=radio name="apply" value="1" <%if apply=1 then%>checked<%end if%> id=apply_1><label for=apply_1>开放</label>
</td>
</tr>


  <tr class=a4>
<td width="50%">类别论坛与下属论坛拥有同样功能：
（如：<font color="FF0000">发帖、浏览</font>）</td>
<td>
<input type=radio name="SortShowForum" value="0" <%if SortShowForum=0 then%>checked<%end if%> id=SortShowForum><label for=SortShowForum>关闭</label>
<input type=radio name="SortShowForum" value="1" <%if SortShowForum=1 then%>checked<%end if%> id=SortShowForum_1><label for=SortShowForum_1>开放</label>
</td></tr>


<tr class=a3>
<td width="50%">打开来源检测（防止从外部提交数据）：</td>
<td>
<input type=radio name="StopOutPost" value="0" <%if StopOutPost=0 then%>checked<%end if%> id=StopOutPost><label for=StopOutPost>关闭</label>
<input type=radio name="StopOutPost" value="1" <%if StopOutPost=1 then%>checked<%end if%> id=StopOutPost_1><label for=StopOutPost_1>打开</label>
</td>
</tr>



<tr class=a4>
<td>首页显示论坛列表样式：</td>
<td>
<input type=radio name="ListStyle" value="0" <%if ListStyle=0 then%>checked<%end if%> id=ListStyle><label for=ListStyle>简洁</label>
<input type=radio name="ListStyle" value="1" <%if ListStyle=1 then%>checked<%end if%> id=ListStyle_1><label for=ListStyle_1>详细</label>
</td>
</tr>




</table>

<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>上传选项</td>
  </tr>
  
<tr class=a4>
<td width="50%">选择上传组件：</td>
<td>
<select name="selectup">
<option value="">关闭</option>
<option value="FSO" <%if selectup="FSO" then%>selected<%end if%>>FSO</option>
<option value="SA-FileUp" <%if selectup="SA-FileUp" then%>selected<%end if%>>SA-FileUp</option>
</select></td>
</tr>
  
<tr class=a3>
<td width="50%">允许头像文件的大小：</td>
<td>
<input size="6" value="<%=MaxFace%>" name="MaxFace" readonly> byte&nbsp; [默认:10240]</td>
</tr>

<tr class=a4>
<td>允许照片文件的大小：</td>
<td>
<input size="6" value="<%=MaxPhoto%>" name="MaxPhoto" readonly> byte&nbsp; [默认:30720]</td>
</tr>
  
 <tr class=a3>
<td>允许附件文件的大小：</td>
<td>
<input size="6" value="<%=MaxFile%>" name="MaxFile" readonly> byte&nbsp; [默认:102400]</td>
</tr>
   
  
   <tr class=a4>
<td>允许附件文件的类型：<br>例如：<font color="FF0000">gif|jpg|png|bmp|swf|txt|mid|doc|xls|zip|rar</font></td>
<td>
<input size="40" value="<%=UpFileGenre%>" name="UpFileGenre" readonly> </td>
</tr>

</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25 class=a1>
    <td class=a4 align=middle colspan=2>过滤设置</td>
  </tr>

  <tr class=a4>
<td width="50%">过滤敏感字（多字请用“|”分隔）：<br>例如：<font color="FF0000">fuck|shit|你妈</font></td>
<td><input size="40" value="<%=badwords%>" name="badwords"></td>
</tr>

  <tr class=a3>
<td width="50%">过滤用户帖子（多用户请用“|”分隔）：<br>例如：<font color="FF0000">yuzi|裕裕</font></td>
<td><input size="40" value="<%=badlist%>" name="badlist"></td>
</tr>

<tr class=a4>
<td width="50%">禁止IP地址进入论坛（多IP请用“|”分隔）：<br>例如：<font color="FF0000">127.0.0.1|192.168.0.1</font></td>
<td><input size="40" value="<%=badip%>" name="badip"></td>
</tr>

</table>


<br>






<input type="submit" value=" 更 新 "></form>
<%
end sub


sub variableok

if Request("clubname")="" then error2("请输入社区名称")
if Request("style")="" then error2("默认风格不能没有设置")
if Request("SendPassword")="1" and Request("selectmail")="" then error2("您选择了注册用户密码通过Email发送，但是您没有设定发送邮件组件")


filtrate=split(Request("badip"),"|")
for i = 0 to ubound(filtrate)
if instr("|"&Request.ServerVariables("REMOTE_ADDR")&"","|"&filtrate(i)&"") > 0 then error2("请检查您输入的IP地址是否正确")
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
If -2147221005 = Err Then error2("本服务器不支持 JMail.Message 组件！请关闭发送邮件功能！")
elseif Request("selectmail")="CDONTS" then
Set MailObject = Server.CreateObject("CDONTS.NewMail")
If -2147221005 = Err Then error2("本服务器不支持 CDONTS.NewMail 组件！请关闭发送邮件功能！")
end if


%>
更新成功<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub



htmlend

%>