<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<script>
function delclass(classid)
{
if (confirm('删除此一级分类的同时将删除该一级分类下的所有二级分类\n以及二级分类下所有舞曲！\n此操作删除的数据将无法挽回，确定要删除此一级分类吗？'))
window.open('ClassSave.asp?act=del&Classid='+classid,'_self')
}
  </script>
<%
dim i
set rs=server.createobject("adodb.recordset")
sql="select * from class order by classid"
rs.open sql,conn,1,1
i=request("id")
if i="" then i=rs("Classid")
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




      <div align="center">
        <center>
      <table border="0" width="100%" cellspacing="3" cellpadding="3" style="border-collapse: collapse">
        <tr>
          <td width="100%" height="20" align=center bgcolor="#FFCB7D"><b>舞 曲 分 类 列 表</b></td>
        </tr>
		<tr>
          <td width="100%" height="20" align=center bgcolor="#EA8C00">
<%do while not rs.eof
  %>
  <a href=classmana.asp?id=<%=rs("Classid")%>><%=rs("Class")%></a>
  <%rs.movenext
	loop
	rs.close
	%>		  
		  </td>
        </tr>
<%sql="select * from class where classid="&i
rs.open sql,conn,1,1
%>
        <tr>
         <td width="100%" height=30>
            <table border="0" cellspacing="3" cellpadding="3" bordercolor="#333333" style="border-collapse: collapse" width="50%">
          <form method="POST" action="ClassSave.asp?act=rename&Classid=<%=rs("Classid")%>" id=form<%=i%> name=form<%=i%>>
              <tr>
                <td><input size=10 type="text" name="Class" value="<%=rs("Class")%>"><input type=button onclick="javascript:submit();" title="修改一级栏目名字" value="改名">
                <input type=button onclick=javascript:delclass("<%=rs("classid")%>") title="删除此一级栏目" value="删除">
                </td>
              </tr>
            </form>
<%
rs.close
%>

            </table>

      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bgcolor="#FFCB7D">
      <form method="POST" action="ClassSave.asp?act=add">
        <tr>
          <td width="100%" height="20" colspan=2 align=center><b>添 加 
          一 级 分 类</b></td>
        </tr>
        <tr align="center">
          <td>类名：<input type="text" name="Class" size="20"> <input type="submit" value="添加" name="B3"></td>
        </tr>
        </table>
      
          </td>
          </form>
        </tr>
      

      </table>
        </center>
      </div>
      
      </td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>