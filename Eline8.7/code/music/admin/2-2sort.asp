<%PageName="1-1sort"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<%
dim i
set rs=server.createobject("adodb.recordset")
set rs2=server.createobject("adodb.recordset")

sql="select * from class order by classid"
rs.open sql,conn,1,1
i=request("id")
if i="" then i=rs("Classid")
%>
<%do while not rs.eof
 %>
<%rs.movenext
loop
rs.close
%>	
<%sql="select * from class where classid="&i
rs.open sql,conn,1,1
if rs.eof then%>
<%
else
%>
     <div align="center">
      <center>
      <table border="1" width="90%" cellspacing="0" cellpadding="0" class="TableLine" bordercolorlight="#CC0066">
        <tr>
          <td width="100%" height="20" bgcolor="#CC0066" align=center><font color="white"><b>二 级 分 类 管 理</b></font></td>
        </tr>
<%
		sql2="select * from SClass where Classid="&rs("Classid")&" order by Sclassid"
		rs2.open sql2,conn,1,1
		if rs2.eof then
%>
              <tr><td align="center">尚无任何二级分类</td></tr>
<%
		else
			j=0
			do while not rs2.eof
			j=j+1
%>

		<tr>
          <td width="100%" height="25" align=left>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <form method="POST" action="SClassSave.asp?act=rename&SClassid=<%=rs2("SClassid")%>" id=form<%=j%> name=form<%=j%>>
                    <tr>
                      <td width="30%">&nbsp;&nbsp;<%=rs2("SClass")%></td>
                      <td width="33%"><input size=15 type="text" name="SClass" value="<%=rs2("SClass")%>">&nbsp;&nbsp;<input style="color: #FFFFFF; background-color: #FF1171; border: 1px solid #000000" type="submit" value="改 名" name="Submit"></td>
                      <td width="37%" align="right"><a title="慎重哦!" href='SClassSave.asp?act=del&SClassid=<%=rs2("SClassID")%>'>删除</a></td>
                    </tr>
                   </form>
                  </table>
                </div>
		  </td>
        </tr>
<%
			rs2.movenext
			loop
		end if
		rs2.close
%>
<%
end if
rs.close
%>
      </table>
     </center>
    </div>
<% 
set rs=nothing
set rs2=nothing
conn.close
set conn=nothing
%></body></html>


