<%PageName="Songadd"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<div align="center">
<center>
<table border="0" width="90%" cellspacing="1" cellpadding="1">
  <tr>
    <td align=center valign=top>
      <table border="1" width="100%" cellspacing="0" cellpadding="2" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <form method="POST" action="SongSave.asp">
          <tr>
           <td width="100%" height="20" colspan=2 bgcolor="#FF1171" align=center><font color="white"><b>添 加 歌 曲</b></font></td>
          </tr>
          <tr>
            <td align="right">* 类型：</td>
            <td>一级栏目:
              <select name="classid" size="1" onchange="window.open('Songadd.asp?classid='+this.options[this.selectedIndex].value,'_self')">
                <option value="" <%if request("classid")="" then%> selected<%end if%>>选择栏目</option>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from class"
rs.open sql,conn,1,1
do while not rs.eof
%>
                <option<%if cstr(request("classid"))=cstr(rs("classid")) and request("classid")<>"" then%> selected<%end if%> value="<%=CStr(rs("classID"))%>" name=classid><%=rs("class")%></option>
<%
rs.movenext
loop
rs.close
%>
              </select>
              二级栏目:
<%if request("classid")<>"" then%>
              <select name="sclassid" size="1" onchange="window.open('Songadd.asp?classid=<%=request("classid")%>&sclassid='+this.options[this.selectedIndex].value,'_self')">
                <option value="" <%if request("sclassid")="" then%> selected<%end if%>>选择栏目</option>
<%
	sql="select * from sclass where classid="&request("classid")
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if cstr(request("sclassid"))=cstr(rs("sclassid")) and request("sclassid")<>"" then%> selected<%end if%> value="<%=CStr(rs("sclassid"))%>" name=sclassid><%=rs("sclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close
%>
<%else%>
              <select name="sclassid" size="1">
                <option value="" selected>选择栏目</option>
<%end if%>
              </select>
              三级栏目:
<%if request("Sclassid")<>"" then%>
              <select name="Nclassid" size="1" onchange="window.open('Songadd.asp?classid=<%=request("classid")%>&SClassid=<%=request("SClassid")%>&nclassid='+this.options[this.selectedIndex].value,'_self')">
                <option value="" <%if request("Nclassid")="" then%> selected<%end if%>>选择栏目</option>
<%
	sql="select * from Nclass where Sclassid="&request("Sclassid")
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if cstr(request("Nclassid"))=cstr(rs("Nclassid")) and request("Nclassid")<>"" then%> selected<%end if%> value="<%=CStr(rs("Nclassid"))%>" name=Nclassid><%=rs("Nclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close
%>
<%else%>
              <select name="Nclassid" size="1">
                <option value="" selected>选择栏目</option>
<%end if%>
              </select>
            </td>
          </tr>
          <tr>
            <td width="15%" align="right">加入专辑：</td>
            <td>
<%if request("Nclassid")<>"" then%>
              <select name="Specialid" size="1">
                <option value="" <%if request("Specialid")="" then%> selected<%end if%>>选择专辑</option>
<%
	sql="select * from special where Nclassid="&request("Nclassid")
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if cstr(request("Specialid"))=CStr(rs("Specialid")) then%> selected<%end if%> value="<%=CStr(rs("Specialid"))%>" name=Specialid><%=rs("name")%></option>
<%
	rs.MoveNext
	Loop
	rs.close
else
%>
              <select name="Specialid" size="1">
                <option value="" selected>选择专辑</option>
<%end if%>
              </select>
            </td>
          </tr>
          <tr>
            <td align="right"><font color="red">歌曲名：</font></td>
            <td><input type="text" name="MusicName" size="20"></td>
          </tr>
          <tr>
            <td width="15%" align="right">Wma地址：</td>
            <td width="85%"><input type="text" name="Wma" size="30"></td>
          </tr>
          <tr>
            <td colspan=2 align=center>
              <input type="hidden" value="add" name="act">
<%if request("Classid")="" then%>
              <input type="button" value=" 确 定 " name="cmdok" onclick="window.alert('请先选择一级栏目！')">&nbsp; 
<%elseif request("SClassid")="" then%>
              <input type="button" value=" 确 定 " name="cmdok" onclick="window.alert('请先选择二级栏目！')">&nbsp; 
<%elseif request("NClassid")="" then%>
              <input type="button" value=" 确 定 " name="cmdok" onclick="window.alert('请先选择三级栏目！')">&nbsp; 
<%else%>
              <input type="submit" value=" 确 定 " name="cmdok">&nbsp; 
<%end if%>
              <input type="reset" value=" 清 除 "  name="cmdcancel">
            </td>
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


