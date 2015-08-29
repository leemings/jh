<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<%
if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" id="AutoNumber2">
    <tr>
      <td width="100%"><img border="0" src="补间.gif" width="1" height="3"></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse" height="155">
  <tr>
    <td valign=top width=150 height="155">
    <!--#include file="admin_left.asp"-->
    　</td>
    <td valign=top width=10 height="155">　</td>
    <td valign=top width="600" height="155">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3" height="100%" bordercolor="#000000">
      <tr>
        <td width="100%" bgcolor="#FF9933" valign="top" height="84">
    <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" width="100%" id="AutoNumber3">
      <tr>
        <td width="100%">
    <%
set rs=server.createobject("adodb.recordset")
sql="select * from user where username<>'短信精灵' order by id desc" 
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
	response.write "<p align='center'>暂时没有用户注册</p>" 
else 
	MaxPerPage=20
	PageUrl="UserMana.asp"
	totalPut=rs.recordcount 
	if currentpage<1 then currentpage=1
	if (currentpage-1)*MaxPerPage>totalput then 
		if (totalPut mod MaxPerPage)=0 then 
			currentpage= totalPut \ MaxPerPage 
		else 
			currentpage= totalPut \ MaxPerPage + 1 
		end if 
	end if 
	if currentPage=1 then 
		showpage totalput,MaxPerPage,PageUrl
		showContent 
		showpage totalput,MaxPerPage,PageUrl
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move  (currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark 
			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		else 
			currentPage=1 
			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		end if 
	end if 
end if 
rs.close 
			
sub showContent 
i=0 
%> 　<div align="center">
      <center>
            <table border="1" width="586" cellspacing="0"  bordercolor="#CC6600" cellpadding="0" style="border-collapse: collapse">
              <tr>
                <td width="25" align=center bgcolor="#F8D7C9">ID</td>
                <td width="226" align=center bgcolor="#F8D7C9">用户名</td>
                <td width="158" align=center bgcolor="#F8D7C9">
                最后一次登陆的IP地址</td>
                <td width="40" align=center bgcolor="#F8D7C9">修改</td>
                <td width="41" align=center bgcolor="#F8D7C9">删除</td>
              </tr>
<%
do while not rs.eof
	i=i+1
%>
              <tr>
                <td align=center width="25"><%=rs("id")%>　</td>
                <td align=center width="226"><a href="UserModify.asp?id=<%=rs("id")%>" title="用户资料：&#13;&#10;密码：<%=rs("Password")%>&#13;&#10;性别：<%if rs("Sex")=true then%>男<%else%>女<%end if%>&#13;&#10;信箱：<%=rs("Email")%>&#13;&#10;OICQ：<%=rs("oicq")%>&#13;&#10;联系地址：<%=rs("Address")%>&#13;&#10;登陆IP：<%=rs("LoginIP")%>&#13;&#10;最后登陆：<%=rs("lastlogin")%>&#13;&#10;注册时间：<%=rs("regDate")%>&#13;&#10;"><%=rs("UserName")%></a>　</td>
                <td align=center width="158"><%=rs("loginip")%>　</td>
                <td align=center width="40"><input type=button name=edit value=修改 onclick="javascript:window.open('UserModify.asp?id=<%=rs("id")%>','_self','')"></td>
                <td align=center width="41"><input type=button name=del value=删除 onclick="javascript:window.open('Usersave.asp?id=<%=rs("id")%>&act=del','_self','')"></td>
              </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
            </table>
      </center>
    </div>
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<form method=Post action="<%=filename%>">
  <center>共<font color="red"><b><%=totalnumber%></b></font>位用户
<%if CurrentPage<2 then%>
  &nbsp;首页 上一页&nbsp;
<%else%>
  &nbsp<a href="<%=filename%>?page=1">首页</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>">上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  下一页 末页
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>">下一页</a>
  <a href="<%=filename%>?page=<%=n%>&classid=<%=classid%>">末页</a>
<%end if%>
  &nbsp;页次:<strong><font color="red"><%=CurrentPage%>/<%=n%></font></strong>页
  转到:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>第<%=i%>页</option>   
<%next%>   
  </select>        
</form>        
<%end function%>
    <div align="center">
      <center>
    
    <table border="1" width="70%" cellspacing="0" bordercolor="#CC6600" cellpadding="0" style="border-collapse: collapse">
       <form method="POST" action="usersave.asp?act=deldate&page=<%=CurrentPage%>" id=form1 name=form1>
         <tr>
           <td width="100%"  align=center bgcolor="#CC6600"><b>删除过期用户</b></td>
         </tr>
         <tr>
           <td width="30%" align=center>最后登陆间隔天数大于：<input type="text" name="selectdate" size="20"></td>
         </tr>
         <tr>
           <td width="30%" align=center>登陆次数小于：<input type="text" name="selecttimes" size="20"></td>
         </tr>
         <tr>
           <td align=center bgcolor="#CC6600">
             <input type="submit" value=" 确定 " name="cmdok">&nbsp; 
             <input type="reset" value=" 清 除 "  name="cmdcancel">
           </td>
         </tr>
         </form>
     </table>

      </center>
    </div>

</td>
      </tr>
    </table>

</td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>