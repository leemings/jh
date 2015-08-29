<%PageName="1-1sort"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<%
dim i
i=0 
MaxList=80
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
          <td width="100%" height="20" bgcolor="#CC0066" align=center><font color="white"><b>一 级 分 类 管 理</b></font></td>
        </tr>
<%
do while not rs.eof
	i=i+1
%>  
		<tr>
          <td width="100%" height="25" align=left>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                   <form method="POST" action="ClassSave.asp?act=rename&Classid=<%=rs("Classid")%>" id=form1 name=form1>
                    <tr>
                      <td width="30%">&nbsp;&nbsp;<%=rs("Class")%></td>
                      <td width="33%"><input size=15 type="text" name="Class" value="<%=rs("Class")%>">&nbsp;&nbsp;<input style="color: #FFFFFF; background-color: #FF1171; border: 1px solid #000000" type="submit" value="改 名" name="Submit"></td>
                      <td width="37%" align="right"><a title="慎重哦!" href='ClassSave.asp?act=del&Classid=<%=rs("ClassID")%>'>删除</a></td>
                    </tr>
                   </form>
                  </table>
                </div>
		  </td>
        </tr>
<%
	if (i mod (MaxList/2)=0) and i>=(MaxList/2) then
%>

<%
	end if
	if i>=MaxList then exit do
    rs.movenext
	loop
	rs.close
	%>
        <tr>
          <td width="100%" height="20" bgcolor="#CC0066" align=center><font color="white"><b>添 加 一 级 分 类</b></font></td>
        </tr>
        <form method="POST" action="ClassSave.asp?act=add">
        <tr>
          <td width="100%" height="20" align=center>类名：<input type="text" name="Class" size="20"> <input type="submit" value="添加" name="B3"></td>
        </tr>
        </form>
      </table>
     </center>
    </div>
<% 
set rs=nothing
conn.close
set conn=nothing
%></body></html>


