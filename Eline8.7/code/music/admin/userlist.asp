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
	response.write "<p align='center'>��ʱû���û�ע��</p>" 
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
            <td width="25%" height=25 align=center bgcolor="#FF1171"><font color="white">�û���</font></td>
            <td width="25%" height=25 align=center bgcolor="#FF1171">email</td>
            <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">�޸�</font></td>
            <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">����</font></td>
            <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">ɾ��</font></td>
          </tr>
          <%
do while not rs.eof
	i=i+1
%>
          <tr> 
            <td height="22" align=center><%=rs("id")%>��</td>
            <td align=center><a href="UserModify.asp?id=<%=rs("id")%>"><%=rs("UserName")%></a></td>
            <td align=center><a href="UserModify.asp?id=<%=rs("id")%>"><%=rs("email")%></a></td>
            <td align=center> 
              <input type=button name=edit value=�޸� onclick="javascript:window.open('UserModify.asp?id=<%=rs("id")%>','_self','')">
            </td>
            <td align=center> 
              <input type=button name=lock value=<%if rs("lockuser")=false then%>"����"<%else%>"����"<%end if%> onclick="javascript:window.open('UserSave.asp?id=<%=rs("id")%>&act=lock','_self','')">
            </td>
            <td align=center> 
              <input type=button name=del value=ɾ�� onclick="javascript:window.open('UserDel.asp?id=<%=rs("id")%>','_self','')">
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
  <center>��<font color="red"><b><%=totalnumber%></b></font>���û�
<%if CurrentPage<2 then%>
  &nbsp;��ҳ ��һҳ&nbsp;
<%else%>
  &nbsp<a href="<%=filename%>?page=1&classid=<%=classid%>&Nclassid=<%=Nclassid%>">��ҳ</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">��һҳ</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  ��һҳ ĩҳ
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">��һҳ</a>
  <a href="<%=filename%>?page=<%=n%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">ĩҳ</a>
<%end if%>
  &nbsp;ҳ��:<strong><font color="red"><%=CurrentPage%>/<%=n%></font></strong>ҳ
  ת��:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
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








