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
          <td width="100%" height="20" bgcolor="#CC0066" align=center><font color="white"><b>歌 手 管 理</b></font></td>
        </tr>
        <tr>
          <td width="100%" height="22" bgcolor="#FFAAD5" align=center>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                   <tr>
                      <td width="20%">&nbsp;&nbsp;<b>↓所属音乐分类</b></td>
                      <td width="40%"><b>↓选择进入歌手列表进行管理</b></td>
                      <td width="40%"><b>↓添加歌手&nbsp;歌手姓名&nbsp;字母索引</b></td>
                    </tr>
                  </table>
                </div>
           </td>
        </tr>
<%
		sql2="select * from SClass where Classid="&rs("Classid")&" order by Sclassid"
		rs2.open sql2,conn,1,1
		if rs2.eof then
%>
<%
		else
			j=0
			do while not rs2.eof
			j=j+1
%>

		<tr>
          <td width="100%" height="29" align=left>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <form method="POST" action="NClassSave.asp?act=add&Classid=<%=rs("Classid")%>&SClassid=<%=rs2("SClassid")%>" align="center">
                    <tr>
                      <td width="20%">&nbsp;&nbsp;<a href=3-3sort.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs2("SClassid")%>><%=rs2("SClass")%></a></td>
                      <td width="40%"><a href=3-3sort.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs2("SClassid")%>><---进入管理此类所属的歌手</a></td>
                      <td width="40%"><input size=13 type="text" name="NClass">&nbsp;&nbsp;<input size=6 type="text" name="Abcd">&nbsp;&nbsp;<input style="color: #FFFFFF; background-color: #FF1171; border: 1px solid #000000" type="submit" value="添 加" name="B1"></td>
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


