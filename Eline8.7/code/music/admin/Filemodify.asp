<%PageName="Filemodify"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<%
Classid=cstr(trim(request.QueryString("classid")))
SClassid=cstr(trim(request.QueryString("sclassid")))
NClassid=cstr(trim(request.QueryString("Nclassid")))
%>
<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<!--#include file="top.asp"-->
<%
Specialid=request.QueryString("Specialid")
AskClassid=request.QueryString("AskClassid")
AskSClassid=request.QueryString("AskSClassid")
AskNClassid=request.QueryString("AskNClassid")
set rs=server.createobject("adodb.recordset")
sql="select * from Special where Specialid="&Specialid
rs.open sql,conn,1,1
if rs.eof then
	errmsg="<li>操作错误！请联系管理员</li>"
	call error()
	Response.End 
else
	name=rs("name")
	classid=rs("classid")
	SClassid=rs("SClassid")
	Nclassid=rs("Nclassid")
	name=rs("name")
	pic=rs("pic")
	Times=rs("Times")
	Yuyan=rs("Yuyan")
	Gongsi=rs("Gongsi")
	intro=rs("intro")
rs.close
end if
%>
     <div align="center">
      <center>
      <table border="1" width="99%" cellspacing="0" cellpadding="0" class="TableLine" bordercolorlight="#CC0066">
        <tr>
          <td width="100%" height="20" bgcolor="#CC0066" align=center><font color="white"><b>修 改 专 辑</b></font></td>
        </tr>
        <tr>
          <td width="100%" height="22" bgcolor="#FFAAD5" align=center>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <form method="POST" action="AddfileSave2.asp?Specialid=<%=Specialid%>&AskClassid=<%=request("AskClassid")%>&AskSClassid=<%=request("AskSClassid")%>&AskNClassid=<%=request("AskNClassid")%>&page=<%=request("page")%>" id=form2 name=form2>
                    <tr>
                      <td width="15%" align="right">类型：</td>
                      <td width="85%">&nbsp;一级栏目：
	          <select name="Classid" size="1" onchange="window.open('FileModify.asp?Specialid=<%=Specialid%>&AskClassid=<%=Askclassid%>&AskSClassid=<%=AskSclassid%>&AskNClassid=<%=AskNclassid%>&page=<%=request("page")%>&Classid='+this.options[this.selectedIndex].value,'_self')">
<%
sql="select * from class"
rs.open sql,conn,1,1
Do until rs.eof
%>
                <option<%if (cstr(request("classid"))=cstr(rs("classid")) and request("classid")<>"") or (cstr(Classid)=cstr(rs("Classid")) and request("classid")="") then%> Selected<%end if%> value="<%=Cstr(rs("classid"))%>"><%=rs("class")%></option>
<%rs.MoveNext
Loop
rs.close
%>            </select>

              &nbsp;二级栏目：
<%if request("classid")<>"" or request("Sclassid")<>"" then
	if request("classid")<>"" then%>
              <select name="Sclassid" size="1" onchange="window.open('FileModify.asp?Specialid=<%=request("Specialid")%>&AskClassid=<%=request("Askclassid")%>&AskSClassid=<%=request("AskSclassid")%>&AskNClassid=<%=request("AskNclassid")%>&page=<%=request("page")%>&Classid=<%=request("Classid")%>&SClassid='+this.options[this.selectedIndex].value,'_self')">
<%	else%>
              <select name="Sclassid" size="1" onchange="window.open('FileModify.asp?Specialid=<%=request("Specialid")%>&AskClassid=<%=request("Askclassid")%>&AskSClassid=<%=request("AskSclassid")%>&AskNClassid=<%=request("AskNclassid")%>&page=<%=request("page")%>&Classid=<%=Classid%>&SClassid='+this.options[this.selectedIndex].value,'_self')">
<%	end if%>
                <option value="" <%if request("Sclassid")="" then%> selected<%end if%>>选择栏目</option>
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
              <select name="Sclassid" size="1" onchange="window.open('FileModify.asp?Specialid=<%=request("Specialid")%>&AskClassid=<%=request("Askclassid")%>&AskSClassid=<%=request("AskSclassid")%>&AskNClassid=<%=request("AskNclassid")%>&page=<%=request("page")%>&Classid=<%=request("Classid")%>&SClassid='+this.options[this.selectedIndex].value,'_self')">
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
            </select>

              &nbsp;三级栏目：
<%if request("SClassid")<>"" or request("Nclassid")<>"" then
	if request("Sclassid")<>"" then%>
              <select name="Nclassid" size="1" onchange="window.open('FileModify.asp?Specialid=<%=request("Specialid")%>&AskClassid=<%=request("Askclassid")%>&AskSClassid=<%=request("AskSclassid")%>&AskNClassid=<%=request("AskNclassid")%>&page=<%=request("page")%>&Classid=<%=request("Classid")%>&SClassid=<%=request("SClassid")%>&NClassid='+this.options[this.selectedIndex].value,'_self')">
<%	else%>
              <select name="Nclassid" size="1" onchange="window.open('FileModify.asp?Specialid=<%=request("Specialid")%>&AskClassid=<%=request("Askclassid")%>&AskSClassid=<%=request("AskSclassid")%>&AskNClassid=<%=request("AskNclassid")%>&page=<%=request("page")%>&Classid=<%=Classid%>&SClassid=<%=request("SClassid")%>&NClassid='+this.options[this.selectedIndex].value,'_self')">
<%	end if%>
                <option value="" <%if request("Nclassid")="" then%> selected<%end if%>>选择栏目</option>
<%
	if request("Sclassid")<>"" then
		sql="select * from Nclass where Sclassid="&request("Sclassid")
	else
		sql="select * from Nclass where Sclassid="&Sclassid
	end if
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if (cstr(request("Nclassid"))=cstr(rs("Nclassid")) and request("Nclassid")<>"") then%> selected<%end if%> value="<%=CStr(rs("Nclassid"))%>" name=Nclassid><%=rs("Nclass")%>ff</option>
<%
	rs.MoveNext
	Loop
	rs.close

else
%>
              <select name="Nclassid" size="1" onchange="window.open('FileModify.asp?Specialid=<%=request("Specialid")%>&AskClassid=<%=request("Askclassid")%>&AskNClassid=<%=request("AskNclassid")%>&page=<%=request("page")%>&Classid=<%=request("Classid")%>&SClassid=<%=request("SClassid")%>&NClassid='+this.options[this.selectedIndex].value,'_self')">
                <option value="">选择栏目</option>
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
            </select></td>
                    </tr>
                  </table>
                </div>
           </td>
        </tr>
		<tr>
          <td width="100%" height="25" align=left>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                      <td width="15%" align="right">专辑名称：</td>
                      <td width="85%"><input type="text" name="name" value="<%=name%>" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right">所属语言：</td>
                      <td width="85%"><input type="text" name="Yuyan" value="<%=Yuyan%>" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right">唱片公司：</td>
                      <td width="85%"><input type="text" name="Gongsi" value="<%=Gongsi%>" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right">发行日期：</td>
                      <td width="85%"><input type="text" name="Times" value="<%=Times%>" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right">专辑图片：</td>
                      <td width="85%"><input type="text" name="pic" value="<%=pic%>" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right">专辑简介：</td>
                      <td width="85%"><TEXTAREA name="intro" rows=5 cols="60"><%=htmlencode1(intro)%></TEXTAREA></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right"></td>
                      <td width="85%"></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right"></td>
                      <td width="85%"></td>
                    </tr>

                  </table>
                </div>
		  </td>
        </tr>
        <tr>
          <td width="100%" height="22" bgcolor="#FFAAD5" align=center>
              <input type="hidden" value="edit" name="act"> 
              <input type="submit" value=" 修 改 " name="cmdok2">&nbsp; 
              <input type="reset" value=" 清 除 " name="cmdcance2l"></td>
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