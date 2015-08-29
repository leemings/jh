<%PageName="SongModify"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<!--#include file="Star.inc"-->
<%
id=request.QueryString("id")
AskClassid=request.QueryString("AskClassid")
AskSClassid=request.QueryString("AskSClassid")
AskNClassid=request.QueryString("AskNClassid")
page=request.QueryString("page")
set rs=server.createobject("adodb.recordset")
sql="select * from MusicList where id="&id
rs.open sql,conn,1,1
if rs.eof then
	errmsg=errmsg+"<br>"+"<li>操作错误！请联系管理员"
	call error()
	Response.End 
else
	Wma=rs("Wma")
	MusicName=rs("MusicName")
	Classid=rs("Classid")
	SClassid=rs("SClassid")
	Nclassid=rs("Nclassid")
	Specialid=rs("Specialid")
	singer=rs("singer")
end if
rs.close
%>
<div align="center">
<center>
<table border="0" width="90%" cellspacing="1" cellpadding="1">
  <tr>
    <td align=center valign=top>
      <table border="1" width="100%" cellspacing="0" cellpadding="2" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <form method="POST" action="SongSave.asp?id=<%=id%>&AskClassid=<%=AskClassid%>&AskSClassid=<%=AskSClassid%>&AskNClassid=<%=AskNClassid%>&page=<%=page%>">
          <tr>
           <td width="100%" height="20" colspan=2 bgcolor="#FF1171" align=center><font color="white"><b>修 改 歌 曲</b></font></td>
          </tr>
          <tr>
            <td align="right" valign="top">类型：</td>
            <td>&nbsp;&nbsp;
              <select class="smallSel" name="classid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid='+this.options[this.selectedIndex].value,'_self')">
<%
set rs=server.createobject("adodb.recordset")
sql="select * from class"
rs.open sql,conn,1,1
do while not rs.eof
%>
                <option<%if (cstr(request("classid"))=cstr(rs("classid")) and request("classid")<>"") or (cstr(Classid)=cstr(rs("Classid")) and request("classid")="") then%> selected<%end if%> value="<%=CStr(rs("classID"))%>" name=classid><%=rs("class")%></option>
<%
rs.movenext
loop
rs.close
%>
              </select>&nbsp;&nbsp;
<%if request("classid")<>"" or request("Sclassid")<>"" then
	if request("classid")<>"" then%>
              <select name="Sclassid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=request("Classid")%>&SClassid='+this.options[this.selectedIndex].value,'_self')">
<%	else%>
              <select name="Sclassid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=Classid%>&SClassid='+this.options[this.selectedIndex].value,'_self')">
<%	end if%>
                <option value="" <%if request("Nclassid")="" then%> selected<%end if%>>选择栏目</option>
<%
	if request("classid")<>"" then
		sql="select * from Sclass where classid="&request("classid")
	else
		sql="select * from Sclass where classid="&classid
	end if
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if (cstr(request("Sclassid"))=cstr(rs("Sclassid")) and request("Sclassid")<>"") then%> selected<%end if%> value="<%=CStr(rs("Sclassid"))%>" name=Sclassid><%=rs("Sclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close

else
%>
              <select name="Sclassid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=request("Classid")%>&SClassid='+this.options[this.selectedIndex].value,'_self')">
                <option value="">选择栏目</option>
<%
	sql="select * from Sclass where classid="&Classid
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if cstr(Sclassid)=cstr(rs("Sclassid")) then%> selected<%end if%> value="<%=CStr(rs("Sclassid"))%>" name=Sclassid><%=rs("Sclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close

end if
%>
              </select>&nbsp;&nbsp;  
<%if request("Sclassid")<>"" or request("Nclassid")<>"" then
	if request("Sclassid")<>"" then%>
              <select name="Nclassid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=request("Classid")%>&SClassid=<%=request("SClassid")%>&NClassid='+this.options[this.selectedIndex].value,'_self')">
<%	else%>
              <select name="Nclassid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=Classid%>&SClassid=<%=SClassid%>&NClassid='+this.options[this.selectedIndex].value,'_self')">
<%	end if%>
                <option value="" <%if request("Nclassid")="" then%> selected<%end if%>>选择子栏目</option>
<%
	if request("Sclassid")<>"" then
		sql="select * from Nclass where Sclassid="&request("Sclassid")
	else
		sql="select * from Nclass where Sclassid="&Sclassid
	end if
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if (cstr(request("Nclassid"))=cstr(rs("Nclassid")) and request("Nclassid")<>"") then%> selected<%end if%> value="<%=CStr(rs("Nclassid"))%>" name=Nclassid><%=rs("Nclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close

else
%>
              <select name="Nclassid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=request("Classid")%>&SClassid=<%=request("SClassid")%>&NClassid='+this.options[this.selectedIndex].value,'_self')">
                <option value="">选择子栏目</option>
<%
	sql="select * from Nclass where Sclassid="&SClassid
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if cstr(Nclassid)=cstr(rs("Nclassid")) then%> selected<%end if%> value="<%=CStr(rs("Nclassid"))%>" name=Nclassid><%=rs("Nclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close

end if
%>
              </select>
            </td>
          </tr>
          <tr>
            <td align=right>专辑：</td>
            <td>
<%
if request("Classid")<>"" or request("Nclassid")<>"" or request("Specialid")<>"" then
	if request("Classid")<>"" then
%>
              <select name="Specialid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=request("Classid")%>&SClassid=<%=request("SClassid")%>&NClassid=<%=request("NClassid")%>&Specialid='+this.options[this.selectedIndex].value,'_self')">
<%	elseif request("NClassid")<>"" then%>
              <select name="Specialid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=Classid%>&SClassid=<%=request("SClassid")%>&NClassid=<%=request("NClassid")%>&Specialid='+this.options[this.selectedIndex].value,'_self')">
<%	else%>
              <select name="Specialid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=Classid%>&SClassid=<%=request("SClassid")%>&NClassid=<%=NClassid%>&Specialid='+this.options[this.selectedIndex].value,'_self')">
<%	end if%>
                <option value="" <%if request("Specialid")="" then%> selected<%end if%>>选择专辑</option>
<%
	if request("NClassid")<>"" then
		sql="select * from special where Nclassid="&request("Nclassid")
	else
		sql="select * from special where Nclassid="&Nclassid
	end if
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
              <select name="Specialid" size="1" onchange="window.open('SongModify.asp?id=<%=request("id")%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=page%>&Classid=<%=request("Classid")%>&SClassid=<%=request("SClassid")%>&NClassid=<%=request("NClassid")%>&Specialid='+this.options[this.selectedIndex].value,'_self')">
                <option value="">选择专辑</option>
<%
	sql="select * from special where Nclassid="&Nclassid
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if cstr(Specialid)=cstr(rs("Specialid")) then%> selected<%end if%> value="<%=CStr(rs("Specialid"))%>" name=Specialid><%=rs("name")%></option>
<%
	rs.MoveNext
	Loop
	rs.close

end if%>
              </select>
            </td>
          </tr>
          <tr>
            <td align="right">歌曲名：</td>
            <td><input type="text" name="MusicName" size="15" class="smallinput" maxlength="100" value="<%=MusicName%>"></td>
          </tr>
          <tr>
            <td width="15%" align="right">Wma地址：</td>
            <td width="85%"><input type="text" name="Wma" value="<%=Wma%>" size="50"></td>
          </tr>
          <tr>
            <td colspan=2 align=center>
              <input type="hidden" value="edit" name="act">
<%
if request("Classid")<>"" then
	if request("SClassid")="" then
%>
              <input type="button" value=" 确 定 " name="cmdok" onclick="window.alert('请先选择二级栏目！')">&nbsp; 
<%	elseif request("NClassid")="" then%>
              <input type="button" value=" 确 定 " name="cmdok" onclick="window.alert('请先选择三级栏目！')">&nbsp; 
<%	elseif request("Specialid")="" then%>
              <input type="button" value=" 确 定 " name="cmdok" onclick="window.alert('请先选择专辑！')">&nbsp; 
<%
	else
%>
              <input type="submit" value=" 确 定 " name="cmdok">&nbsp; 
<%
	end if
else
%>
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







