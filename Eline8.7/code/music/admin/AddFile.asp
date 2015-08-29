<%PageName="1-1sort"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<%
dim i
i=0 
MaxList=8
set rs=server.createobject("adodb.recordset")

sql="select * from class order by classid"
rs.open sql,conn,1,1
i=request("id")
if i="" then i=rs("Classid")
%>
     <div align="center">
      <center>
      <table border="1" width="90%" cellspacing="0" cellpadding="0" class="TableLine" bordercolorlight="#CC0066">
        <tr>
          <td width="100%" height="20" bgcolor="#CC0066" align=center><a href="Addfile3.asp?classid=1"><font color="white"><b>添 加 编 辑 专 辑 (第一步)&nbsp;&nbsp;&nbsp;点 这 里 可 直 接 添 加 > > ></b></font></a></td>
        </tr>
        <tr>
          <td width="100%" height="22" bgcolor="#FFAAD5" align=center>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                   <form method="POST" action="SClassSave.asp?act=add&Classid=<%=rs("Classid")%>" align="center">
                    <tr>
                      <td width="20%">&nbsp;&nbsp;<b>↓所属大类</b></td>
                      <td width="80%"><b>↓请先所属一级分类</b></td>
                      <td width="0%"><b></td>
                    </tr>
                   </form>
                  </table>
                </div>
           </td>
        </tr>
<%
do while not rs.eof
	i=i+1
%>  
		<tr>
          <td width="100%" height="25" align=left>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                   <form method="POST" action="SClassSave.asp?act=add&Classid=<%=rs("Classid")%>" align="center">
                    <tr>
                      <td width="20%">&nbsp;&nbsp;<a href=AddFile1.asp?id=<%=rs("Classid")%>><%=rs("Class")%></a></td>
                      <td width="80%"><a href=AddFile1.asp?id=<%=rs("Classid")%>><---进入选择进行下一步</a></td>
                      <td width="0%"></td>
                    </tr>
                   </form>
                  </table>
                </div>
		  </td>
        </tr>
<%
	if (i mod (MaxList/1)=0) and i>=(MaxList/1) then
%>

<%
	end if
	if i>=MaxList then exit do
    rs.movenext
	loop
	rs.close
	%>
      </table>
     </center>
    </div>
<% 
set rs=nothing
conn.close
set conn=nothing
%></body></html>


