<%PageName="list"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<%
if not isempty(request.QueryString("page")) and request.QueryString("page")<>"" then
	currentPage=cint(request.QueryString("page"))
else
	currentPage=1
end if
if request.QueryString("classid")<>"" then
	classid=request.QueryString("classid")
else
	classid=""
end if
if request.QueryString("Nclassid")<>"" then
	Nclassid=request.QueryString("Nclassid")
else
	Nclassid=""
end if
%>
<!--#include file="top.asp"-->
<%
set rs=conn.execute("SELECT * FROM class order by classid") 
if not Rs.eof then
	if Left(PageName,5)="admin" and PageName<>"index" then
%>
<%
			do while not rs.eof
%>
<%
			rs.movenext
			loop
		elseif PageName="AddFileList" or PageName="AddFileModify" or PageName="AddFileList" then
%>
<%
			do while not rs.eof
%>
<%
			rs.movenext
			loop
		else
%>
<%
			do while not rs.eof
%>
<%
			rs.movenext
			loop
		end if
	else
%>
<%
		do while not rs.eof
%>
<%
		rs.movenext
		loop
	end if
rs.close
%>
<div align="center">
  <center>
<table border="0" width="96%" cellspacing="1" cellpadding="1">
  <tr>
    <td align=center width="100%" valign=top>
<%
if request.QueryString("Nclassid")<>"" then
	sql="select * from Special where Nclassid="+cstr(Nclassid)+" order by Specialid desc" 
elseif request.QueryString("classid")<>"" then
	sql="select * from Special where classid="+cstr(classid)+" order by Specialid desc" 
else
	sql="select * from Special order by Specialid desc" 
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
	response.write "<p align='center'>��ʱû���κ�ר��</p>" 
else 
	MaxPerPage=MaxSpecialList
	PageUrl="AddFileList.asp"
	totalPut=rs.recordcount 
	if currentpage<1 then currentpage=1
	if (currentpage-1)*MaxPerPage>totalput then 
		if (totalPut mod MaxPerPage)=0 then 
			currentpage= totalPut \ MaxPerPage 
		else 
			currentpage= totalPut \ MaxPerPage + 1 
		end if 
	end if 
	if currentPage=1 then 
		showpage totalput,MaxPerPage,PageUrl
		showContent 
		showpage totalput,MaxPerPage,PageUrl
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move  (currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark 
			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		else 
			currentPage=1 
			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		end if 
	end if 
end if 
rs.close 
			
sub showContent 
i=0 
%>
            <table border="1" width="100%" cellspacing="0" cellpadding="5" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
              <tr>
                <td width="33%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">ר������</font></td>
                <td width="12%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">��������</font></td>
                <td width="11%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">�������</font></td>
                <td width="11%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">��ҳ�Ƽ�</font></td>
                <td width="7%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">�޸�</font></td>
                <td width="7%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">ɾ��</font></td>
              </tr>
<%
do while not rs.eof
	i=i+1
%>
              <tr>
                <td width="33%"><a href="SongList.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs("SClassid")%>&NClassid=<%=rs("NClassid")%>&specialid=<%=rs("specialid")%>"><font color=black><%=rs("name")%></font></a>&nbsp; <a href="SongAdd.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs("SClassid")%>&NClassid=<%=rs("NClassid")%>&specialid=<%=rs("specialid")%>">[���]</a></td>
                <td width="12%" align=center><%=rs("Nclass")%></td>
                <td width="11%" align=center><a href="SongConAdd.asp?specialid=<%=rs("specialid")%>"><FONT color="maroon">Wma</FONT></a></td>
                <td width="11%" align=center><a href="AddfileSave.asp?act=SetIsGood&Specialid=<%=rs("Specialid")%>&Askclassid=<%=classid%>&AskSClassid=<%=SClassid%>&AskNclassid=<%=Nclassid%>&page=<%=CurrentPage%>"><%if rs("IsGood")=true then%>����<%else%>�Ƽ�<%end if%></a></td>
                <td width="7%" align=center><a href="Filemodify.asp?Specialid=<%=rs("Specialid")%>&AskClassid=<%=rs("Classid")%>&AskSClassid=<%=rs("SClassid")%>&AskNClassid=<%=rs("NClassid")%>&page=<%=CurrentPage%>">�޸�</a></td> 
                <td width="7%" align=center><a href="Filedel.asp?Specialid=<%=rs("Specialid")%>&classid=<%=rs("Classid")%>&AskSClassid=<%=rs("SClassid")%>&Nclassid=<%=rs("NClassid")%>&page=<%=CurrentPage%>">ɾ��</a></td> 
              </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
            </table>
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<form method=Post action="<%=filename%>?classid=<%=classid%>&Nclassid=<%=Nclassid%>">
  <center>����<b><%=totalnumber%></b>��ר��
<%if CurrentPage<2 then%>
  &nbsp;��ҳ ��һҳ&nbsp;
<%else%>
  &nbsp<a href="<%=filename%>?page=1&classid=<%=classid%>&Nclassid=<%=Nclassid%>">��ҳ</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">��һҳ</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  ��һҳ ĩҳ
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">��һҳ</a>
  <a href="<%=filename%>?page=<%=n%>&classid=<%=classid%>&Nclassid=<%=Nclassid%>">ĩҳ</a>
<%end if%>
  &nbsp;ҳ��:<strong><%=CurrentPage%>/<%=n%></strong>ҳ
  ת��:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
<%next%>   
  </select>        
</form>        
<%end function%>
    </td>
  </tr>
  </table>
  </center>
  </div>
<%
set rs=nothing
conn.close
set conn=nothing%></body></html>




