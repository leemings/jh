<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<script>
function delsclass(sclassid)
{
if (confirm('删除此二级分类的同时将删除该二级分类下的所有舞曲！\n此操作删除的数据将无法挽回，确定要删除此二级分类吗？'))
window.open('SClassSave.asp?act=del&SClassid='+sclassid,'_self')
}
  </script>
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
          <td width="100%" height="20" align=center bgcolor="#FFCB7D"><b>二级分类管理</b></td>
        </tr>
		<tr>
          <td width="100%" height="20" align=center bgcolor="#EA8C00">
<%
dim i
set rs=server.createobject("adodb.recordset")
sql="select * from Sclass order by Sclassid"
rs.open sql,conn,1,1
if not rs.eof then
do while not rs.eof
classid=rs("classid")
  %>
          <form method="POST" action="SClassSave.asp?act=edit&SClassid=<%=rs("SClassid")%>" id=form<%=rs("SClassid")%> name=form<%=rs("SClassid")%>>
              所属一级分类：<select name="classid" size="1">
<%
set Trs=server.createobject("adodb.recordset")
sql="select * from class"
Trs.open sql,conn,1,1
do while not Trs.eof
%>
                <option<%if cstr(Classid)=cstr(Trs("classid")) then%> selected<%end if%> value="<%=CStr(Trs("classID"))%>" name=classid><%=Trs("class")%></option>
<%
Trs.movenext
loop
Trs.close
%>
              </select>二级分类名：
              
<input size=10 type="text" name="SClass" value="<%=rs("SClass")%>">
<input type=button onclick="javascript:submit();" title="修改二级栏目名字" value="修 改">
<input type=button onclick=javascript:delsclass("<%=rs("SClassID")%>") title="删除此二级栏目" value="删 除">
<input type=button onclick="window.location.href='songadd.asp?SClassid=<%=rs("SClassid")%>'" title="向此分类中添加舞曲" value="添 歌">
           </form>

  <%

  rs.movenext
	loop
else
end if
	rs.close
	%>		  
		  </td>
        </tr>

            </table>

      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bgcolor="#FFCB7D">
      <form method="POST" action="SClassSave.asp?act=add">
        <tr>
          <td width="100%" height="20" colspan=2 align=center><b>添 加 
          二 级 分 类</b></td>
        </tr>
        <tr align="center">
          <td>所属一级分类：<select name="classid" size="1">
<%
set Trs=server.createobject("adodb.recordset")
sql="select * from class"
Trs.open sql,conn,1,1
do while not Trs.eof
%>
                <option value="<%=CStr(Trs("classID"))%>" name=classid><%=Trs("class")%></option>
<%
Trs.movenext
loop
Trs.close
%>
              </select>
              二级分类名：<input type="text" name="SClass" size="20">
              			<input type="submit" value="添加" name="B3">
              			</td>
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