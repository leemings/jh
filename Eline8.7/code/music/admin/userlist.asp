<%PageName="UserList"%>
<!--#include file="function.asp"-->
<%CheckAdmin3%>
<!--#include file="conn.asp"-->
<%
if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
%>
<!--#include file="top.asp"-->
<div align="center">
<center>
<table border="0" width="86%" cellspacing="1" cellpadding="1">
 <tr>
<td align=center valign=top>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from [user] order by id desc" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
	response.write "<p align='center'>暂时没有用户注册</p>" 
else 
	MaxPerPage=20
	PageUrl="Userlist.asp"
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
%>
            
        <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
          <tr> 
            <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">ID</font></td>
            <td width="25%" height=25 align=center bgcolor="#FF1171"><font color="white">用户名</font></td>
            <td width="25%" height=25 align=center bgcolor="#FF1171">email</td>
            <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">修改</font></td>
            <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">锁定</font></td>
            <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">删除</font></td>
          </tr>
          <%
do while not rs.eof
	i=i+1
%>
          <tr> 
            <td height="22" align=center><%=rs("id")%>　</td>
            <td align=center><a href="UserModify.asp?id=<%=rs("id")%>"><%=rs("UserName")%></a></td>
            <td align=center><a href="UserModify.asp?id=<%=rs("id")%>"><%=rs("email")%></a></td>
            <td align=center> 
              <input type=button name=edit value=修改 onclick="javascript:window.open('UserModify.asp?id=<%=rs("id")%>','_self','')">
            </td>
            <td align=center> 
              <input type=button name=lock value=<%if rs("lockuser")=false then%>"锁定"<%else%>"开锁"<%end if%> onclick="javascript:window.open('UserSave.asp?id=<%=rs("id")%>&act=lock','_self','')">
            </td>
            <td align=center> 
              <input type=button name=del value=删除 onclick="javascript:window.open('UserDel.asp?id=<%=rs("id")%>','_self','')">
            </td>
          </tr>
          <%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
        </table>
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<form method=Post action="<%=filename%>?classid=<%=classid%>&Nclassid=<%=Nclassid%>">
  <center>共<font color="red"><b><%=totalnumber%></b></font>个用户
<%if CurrentPage<2 then%>
  &nbsp;首页 上一页&nbsp;
<%else%>
  &nbsp<a href="<%=filename%>?page=1&classid=<%=classid%>&Nclassid=<%=Nclassid%>">首页</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  下一页 末页
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">下一页</a>
  <a href="<%=filename%>?page=<%=n%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">末页</a>
<%end if%>
  &nbsp;页次:<strong><font color="red"><%=CurrentPage%>/<%=n%></font></strong>页
  转到:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>第<%=i%>页</option>   
<%next%>   
  </select>        
</form>        
<%end function%>
    </td>
  </tr>
  </table>
</div>
<%
set rs=nothing
conn.close
set conn=nothing%></body></html>








