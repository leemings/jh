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
      <td width="100%"><img border="0" src="����.gif" width="1" height="3"></td>
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
    ��</td>
    <td valign=top width=10 height="155">��</td>
    <td valign=top width="600" height="155">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3" height="100%" bordercolor="#000000">
      <tr>
        <td width="100%" bgcolor="#FF9933" valign="top" height="84">
    <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" width="100%" id="AutoNumber3">
      <tr>
        <td width="100%">
    <%
set rs=server.createobject("adodb.recordset")
sql="select * from user where username<>'���ž���' order by id desc" 
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
	response.write "<p align='center'>��ʱû���û�ע��</p>" 
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
%> ��<div align="center">
      <center>
            <table border="1" width="586" cellspacing="0"  bordercolor="#CC6600" cellpadding="0" style="border-collapse: collapse">
              <tr>
                <td width="25" align=center bgcolor="#F8D7C9">ID</td>
                <td width="226" align=center bgcolor="#F8D7C9">�û���</td>
                <td width="158" align=center bgcolor="#F8D7C9">
                ���һ�ε�½��IP��ַ</td>
                <td width="40" align=center bgcolor="#F8D7C9">�޸�</td>
                <td width="41" align=center bgcolor="#F8D7C9">ɾ��</td>
              </tr>
<%
do while not rs.eof
	i=i+1
%>
              <tr>
                <td align=center width="25"><%=rs("id")%>��</td>
                <td align=center width="226"><a href="UserModify.asp?id=<%=rs("id")%>" title="�û����ϣ�&#13;&#10;���룺<%=rs("Password")%>&#13;&#10;�Ա�<%if rs("Sex")=true then%>��<%else%>Ů<%end if%>&#13;&#10;���䣺<%=rs("Email")%>&#13;&#10;OICQ��<%=rs("oicq")%>&#13;&#10;��ϵ��ַ��<%=rs("Address")%>&#13;&#10;��½IP��<%=rs("LoginIP")%>&#13;&#10;����½��<%=rs("lastlogin")%>&#13;&#10;ע��ʱ�䣺<%=rs("regDate")%>&#13;&#10;"><%=rs("UserName")%></a>��</td>
                <td align=center width="158"><%=rs("loginip")%>��</td>
                <td align=center width="40"><input type=button name=edit value=�޸� onclick="javascript:window.open('UserModify.asp?id=<%=rs("id")%>','_self','')"></td>
                <td align=center width="41"><input type=button name=del value=ɾ�� onclick="javascript:window.open('Usersave.asp?id=<%=rs("id")%>&act=del','_self','')"></td>
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
  <center>��<font color="red"><b><%=totalnumber%></b></font>λ�û�
<%if CurrentPage<2 then%>
  &nbsp;��ҳ ��һҳ&nbsp;
<%else%>
  &nbsp<a href="<%=filename%>?page=1">��ҳ</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>">��һҳ</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  ��һҳ ĩҳ
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>">��һҳ</a>
  <a href="<%=filename%>?page=<%=n%>&classid=<%=classid%>">ĩҳ</a>
<%end if%>
  &nbsp;ҳ��:<strong><font color="red"><%=CurrentPage%>/<%=n%></font></strong>ҳ
  ת��:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
<%next%>   
  </select>        
</form>        
<%end function%>
    <div align="center">
      <center>
    
    <table border="1" width="70%" cellspacing="0" bordercolor="#CC6600" cellpadding="0" style="border-collapse: collapse">
       <form method="POST" action="usersave.asp?act=deldate&page=<%=CurrentPage%>" id=form1 name=form1>
         <tr>
           <td width="100%"  align=center bgcolor="#CC6600"><b>ɾ�������û�</b></td>
         </tr>
         <tr>
           <td width="30%" align=center>����½����������ڣ�<input type="text" name="selectdate" size="20"></td>
         </tr>
         <tr>
           <td width="30%" align=center>��½����С�ڣ�<input type="text" name="selecttimes" size="20"></td>
         </tr>
         <tr>
           <td align=center bgcolor="#CC6600">
             <input type="submit" value=" ȷ�� " name="cmdok">&nbsp; 
             <input type="reset" value=" �� �� "  name="cmdcancel">
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