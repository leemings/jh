<%PageName="songlist"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<%
if request("page")<>"" then
	currentPage=cint(request("page"))
else
	currentPage=1
end if

if request("classid")<>"" then
	classid=request("classid")
else
	classid=""
end if
if request("Sclassid")<>"" then
	Sclassid=request("Sclassid")
else
	Sclassid=""
end if
if request("Nclassid")<>"" then
	Nclassid=request("Nclassid")
else
	Nclassid=""
end if
if request("Specialid")<>"" then
	Specialid=request("Specialid")
else
	Specialid=""
end if
%>
<!--#include file="top.asp"-->
<div align="center">
<center>
<table border="0" width="96%" cellspacing="1" cellpadding="1">
 <tr>
<td align=center valign=top>
<%
set rs=server.createobject("adodb.recordset")
if classid<>"" then
	if Specialid<>"" then
		sql="select * from MusicList where Specialid="+cstr(Specialid)+" and Nclassid="+cstr(Nclassid)+" and sclassid="+cstr(sclassid)+" and classid="+cstr(classid)+"  order by id desc" 
	elseif Nclassid<>"" then
		sql="select * from MusicList where Nclassid="+cstr(Nclassid)+" and sclassid="+cstr(sclassid)+" and classid="+cstr(classid)+"  order by id desc" 
	elseif SClassid<>"" then
		sql="select * from MusicList where SClassID="+cstr(SClassID)+" and ClassID="+cstr(Classid)+" order by id desc" 
	else
		sql="select * from MusicList where classid="+cstr(classid)+" order by id desc" 
	end if
else
	sql="select * from MusicList order by id desc"
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
	response.write "<p align='center'><b>��ʱû���ռ��κθ���<br><br><a href='javascript:history.go(-1)'>...::: �� �� �� �� :::...</a></b></p>" 
	showList
else 
	MaxPerPage=50
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
		showpage totalput,MaxPerPage,"songlist.asp" 
		showContent 
		showpage totalput,MaxPerPage,"songlist.asp" 
		showList
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move  (currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark 
			showpage totalput,MaxPerPage,"songlist.asp" 
			showContent 
			showpage totalput,MaxPerPage,"songlist.asp" 
			showList
		else 
			currentPage=1 
			showpage totalput,MaxPerPage,"songlist.asp" 
			showContent 
			showpage totalput,MaxPerPage,"songlist.asp" 
			showList
		end if 
	end if 
	rs.close 
end if 

sub showContent 
dim i 
i=0 
%>
       <table border="1" width="100%" cellspacing="0" cellpadding="2" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <tr>
          <td width="40%" height=22 align=center bgcolor="#FF1171"><font color="white">��������</font></td>
          <td width="30%" height=22 align=center bgcolor="#FF1171"><font color="white">��������</font></td>
          <td width="9%" height=22 align=center bgcolor="#FF1171"><font color="white">�Ƽ�</font></td>
          <td width="9%" height=22 align=center bgcolor="#FF1171"><font color="white">�޸�</font></td>
          <td width="9%" height=22 align=center bgcolor="#FF1171"><font color="white">ɾ��</font></td>
        </tr>
<%do while not rs.eof%>
        <tr>
          <td width="40%"> <%=(i+1)%>.<a href="<%=rs("Wma")%>"><%=rs("MusicName")%></a></td>
          <td width="30%" align=center><%=rs("singer")%>  <a href="SongAdd.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs("SClassid")%>&NClassid=<%=rs("NClassid")%>&specialid=<%=rs("specialid")%>">[�������]</a></td>
          <td width="9%" align=center><a href="SongSave2.asp?act=SetIsGood&id=<%=rs("id")%>&classid=<%=classid%>&SClassid=<%=SClassid%>&Nclassid=<%=Nclassid%>&page=<%=CurrentPage%>"><%if rs("IsGood")=true then%>����<%else%>�Ƽ�<%end if%></a></td> 
          <td width="9%" align=center><a href="SongModify.asp?id=<%=rs("id")%>&AskClassid=<%=classid%>&AskNClassid=<%=Nclassid%>&page=<%=CurrentPage%>">�޸�</a></td> 
          <td width="9%" align=center><a href="SongDel.asp?id=<%=rs("id")%>&classid=<%=classid%>&SClassid=<%=SClassid%>&Nclassid=<%=Nclassid%>&page=<%=CurrentPage%>">ɾ��</a></td> 
        </tr>
<%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
	loop
%>
      </table>
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
dim n
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<form method=Post action="<%=filename%>?classid=<%=classid%>&Sclassid=<%=Sclassid%>&Nclassid=<%=Nclassid%>">
  <center>��<font color="#ff0000"><b><%=totalnumber%></b></font>�׸���
<%if CurrentPage<2 then%>
  &nbsp;��ҳ ��һҳ&nbsp;
<%else%>
  &nbsp<a href="<%=filename%>?page=1&classid=<%=classid%>&Sclassid=<%=Sclassid%>&Nclassid=<%=Nclassid%>">��ҳ</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&classid=<%=classid%>&Sclassid=<%=Sclassid%>&Nclassid=<%=Nclassid%>">��һҳ</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  ��һҳ ĩҳ
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&classid=<%=classid%>&Sclassid=<%=Sclassid%>&Nclassid=<%=Nclassid%>">��һҳ</a>
  <a href="<%=filename%>?page=<%=n%>&classid=<%=classid%>&Sclassid=<%=Sclassid%>&Nclassid=<%=Nclassid%>">ĩҳ</a>
<%end if%>
  &nbsp;ҳ��:<strong><font color="#ff0000"><%=CurrentPage%>/<%=n%></font></strong>ҳ
  ת��:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
<%next%>   
  </select>        
</form>        

<% 
end function
%>
<%
sub showList
%>
       <table border="1" width="100%" cellspacing="0" cellpadding="2" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <tr>
          <td colspan=2 width="100%" height=22 align=center bgcolor="#FF1171"><font color="white">��Ŀ�б�</font></td>
        <tr>
<%
'-----------------һ��Ŀ¼�б�-----------------------
set Trs=server.createobject("adodb.recordset")
Tsql="select Classid,Class from Class"
Trs.open Tsql,conn,1,1
if not Trs.eof then
%>
        <tr>
          <td width=12%>һ����Ŀ:</td>
          <td>
<%	do while not Trs.eof%>
            <a href="songlist.asp?Classid=<%=Trs("Classid")%>"><%=Trs("Class")%></a>
<%
	Trs.movenext
	loop
%> ��</td>
        </tr>
<%
end if
Trs.close
'-----------------����Ŀ¼�б�-----------------------
if classid<>"" then
	Tsql="select Classid,SClassid,SClass from SClass where classid="+cstr(classid) 
	Trs.open Tsql,conn,1,1
	if not Trs.eof then
%>
        <tr>
          <td width=12%>������Ŀ:</td>
          <td>
<%		do while not Trs.eof%>
            <a href="songlist.asp?Classid=<%=Trs("Classid")%>&SClassid=<%=Trs("SClassid")%>"><%=Trs("SClass")%></a>
<%
		Trs.movenext
		loop
%> ��</td>
        </tr>
<%
	end if
	Trs.close
end if
'-----------------����Ŀ¼�б�-----------------------
if Sclassid<>"" then
	Tsql="select Classid,SClassid,NClassid,NClass from NClass where sclassid="+cstr(sclassid)
	Trs.open Tsql,conn,1,1
	if not Trs.eof then
%>
        <tr>
          <td width=12%>������Ŀ:</td>
          <td>
<%		do while not Trs.eof%>
            <a href="songlist.asp?Classid=<%=Trs("Classid")%>&SClassid=<%=Trs("SClassid")%>&NClassid=<%=Trs("NClassid")%>"><%=Trs("NClass")%>
<%
		Trs.movenext
		loop
%>
          ��</td>
        </tr>
<%
	end if
	Trs.close
end if
%>
      </td>
    </table>
<%
end sub
%>
    </td>
  </tr>
</table>
</div>
<%
set Trs=nothing
set rs=nothing
conn.close
set conn=nothing
%></body></html>


