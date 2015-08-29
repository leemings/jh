<%PageName="admin_ResearchMana"%>
<!--#include file="function.asp"-->
<%CheckAdmin2%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from Poll order by id desc"
rs.open sql,conn,1,1
%>
<table border="0" width="100%" cellspacing="0" cellpadding="5">
  <tr>
    <td align=center valign=top>
      <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
          <tr>
            <td width="100%" height="20" colspan=5 bgcolor="#FF1171" align=center><font color="white"><b>管 理 首 页 调 查</b></font></td>
          </tr>
          <tr>
            <td width="10%" height="20" align="center" bgcolor="#FFD2E4">ID</td>
            <td width="60%" align="center" bgcolor="#FFD2E4">主题</td>
            <td width="10%" align="center" bgcolor="#FFD2E4">修改</td>
            <td width="10%" align="center" bgcolor="#FFD2E4">删除</td>
          </tr>
<%
do while not rs.eof
%>
          <tr>
            <td width="10%" align="center" height="20"><%=rs("ID")%>　</td>
            <td width="60%"><%=rs("Title")%>　</td>
            <td width="10%" align="center"><input onclick="javascript:window.open('PollModify.asp?id=<%=rs("ID")%>','_self','')" type="button" value="修改" name="button1"></td>
            <td width="10%" align="center"><input onclick="javascript:window.open('PollDel.asp?id=<%=rs("ID")%>','_self','')" type="button" value="删除" name="button2"></td>
          </tr>
<%
rs.movenext
loop
rs.close
%>
          <tr>
            <td colspan=4 align=center><input onclick="javascript:window.open('PollAdd.asp','_self','')" type="button" value="添加新调查" name="button"></td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
<%
set rs=nothing
conn.close
set conn=nothing%></body></html>


