<%PageName="search"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<%
keyword=trim(request("keyword")) 
keyword=replace(keyword,"'","''") 
stype=request("stype")
if keyword="" then
	errmsg=errmsg+"查找字符不能为空，请重输入查找的信息<a href=""javascript:history.go(-1)"">返回重查</a>"
	call error()
	Response.End 
end if
%>
<!--#include file="top.asp"-->
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="100%" valign="top">
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
          <tr>
            <td width="100%" bgcolor="#FFB5D2" height="19" valign="bottom">&nbsp; 
              <b> 搜 索 结 果 :::...</b></td>
            <td width="0%" bgcolor="#FFB5D2"></td>
          </tr>
          <tr>
            <td width="100%" colspan="2">
<%
'---------------------------search----------------------
Set rs= Server.CreateObject("ADODB.Recordset")
if stype="Singer" then
	sql="select * from Nclass where Nclass Like '%"& keyword &"%' order by Nclassid desc"
elseif stype="Music" then 
	sql="select * from MusicList where MusicName Like '%"& keyword &"%' order by id desc"
elseif stype="Special" then
	sql="select * from Special where Name Like '%"& keyword &"%' order by Specialid desc"
elseif stype="User" then
	sql="select * from [User] where username Like '%"& keyword &"%' order by id desc"
else
end if
rs.open sql,conn,1,1

if not isempty(request.QueryString("page")) then
	currentPage=cint(request.QueryString("page"))
else
	currentPage=1
end if
if rs.eof and rs.bof then
	if stype="Singer" then
		response.write "<p align='center'><br><br>Sorry,  未 找 到 你 想 要 的 歌 手<br><br><a href=""javascript:history.go(-1)"">点 此 返 回</a><br><br></p>"
	elseif stype="Music" then 
		response.write "<p align='center'><br><br>Sorry,  未 找 到 你 想 要 的 歌 曲<br><br><a href=""javascript:history.go(-1)"">点 此 返 回</a><br><br></p>"
    elseif stype="Special" then
		response.write "<p align='center'><br><br>Sorry,  未 找 到 你 想 要 的 专 辑<br><br><a href=""javascript:history.go(-1)"">点 此 返 回</a><br><br></p>"
	elseif stype="User" then
		response.write "<p align='center'><br><br>Sorry,  未 找 到 你 想 找 的 用 户<br><br><a href=""javascript:history.go(-1)"">点 此 返 回</a><br><br></p>"
	end if
else
	totalPut=rs.recordcount
	MaxPerPage=10
	PageUrl="Search.asp"
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
		search
	else
		if (currentPage-1)*MaxPerPage<totalPut then
			rs.move  (currentPage-1)*MaxPerPage
			dim bookmark
			bookmark=rs.bookmark
			showpage totalput,MaxPerPage,PageUrl
			showContent
			showpage totalput,MaxPerPage,PageUrl
			search
		else
			currentPage=1
			showpage totalput,MaxPerPage,PageUrl
			showContent
			showpage totalput,MaxPerPage,PageUrl
			search
		end if
	end if
	rs.close
end if

sub showContent 
i=0 
%>
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
                <td><div align="center">
<!-----------------------搜索歌手----------------------------->
<%if stype="Singer" then%>
                  <table border="1" width="88%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
                    <tr>
                      <td width="35%" height=22 align=center bgcolor="#FF1171"><font color="white">歌手</font></td>
                      <td width="50%" height=22 align=center bgcolor="#FF1171"><font color="white">修改</font></td>
                      <td width="15%" height=22 align=center bgcolor="#FF1171"><font color="white">删除</font></td>
                    </tr>
<%
	Set Trs= Server.CreateObject("ADODB.Recordset")
	do while not rs.eof
		i=i+1
		Tsql="SELECT * FROM MusicList where Nclassid="+cstr(rs("Nclassid"))
		Trs.open Tsql,conn,1,1
		TotalMNum=Trs.recordcount
		Trs.close
		Tsql="SELECT * FROM Special where Nclassid="+cstr(rs("Nclassid"))
		Trs.open Tsql,conn,1,1
		TotalSNum=Trs.recordcount
		Trs.close
		Tsql="SELECT * FROM Nclass where Nclassid="+cstr(rs("Nclassid"))
		Trs.open Tsql,conn,1,1
		Trs.close
%>
                   <form method="POST" action="NClassSave.asp?act=rename&NClassid=<%=rs("NClassid")%>" id=Nform<%=k%> name=Nform<%=k%>>
                    <tr>
                      <td>&nbsp;&nbsp;<%=i%>.&nbsp;<%=rs("Nclass")%></td>
                      <td align=center><input type="text" name="NClass" size="15" value="<%=rs("NClass")%>">&nbsp;&nbsp;<input type="text" name="Abcd" size="6" value="<%=rs("Abcd")%>">&nbsp;&nbsp;<input style="color: #FFFFFF; background-color: #FF1171; border: 1px solid #000000" type="submit" value="改 名" name="submit"></td>
                      <td align=center><a title="慎重! 删除后可是不能恢复了哦!" href='NClassSave.asp?act=del&NClassid=<%=rs("NClassID")%>'>删除</a></td>
                    </tr>
                   </form>
<%
		if i>=MaxPerPage then exit do
	rs.movenext
	loop
	set Trs=nothing
%>
                  </table>

<!-----------------------搜索歌曲----------------------------->
<%elseif stype="Music" then%>
       <table border="1" width="90%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <tr>
          <td width="40%" height=22 align=center bgcolor="#FF1171"><font color="white">歌曲名字</font></td>
          <td width="30%" height=22 align=center bgcolor="#FF1171"><font color="white">所属歌手</font></td>
          <td width="9%" height=22 align=center bgcolor="#FF1171"><font color="white">推荐</font></td>
          <td width="9%" height=22 align=center bgcolor="#FF1171"><font color="white">修改</font></td>
          <td width="9%" height=22 align=center bgcolor="#FF1171"><font color="white">删除</font></td>
        </tr>
<%do while not rs.eof%>
        <tr>
          <td width="40%"> <%=(i+1)%>.<a href="javascript:open_window('song.asp?id=<%=rs("id")%>','Listen','width=250,height=100')"><%=rs("MusicName")%></a></td>
          <td width="30%" align=center><%=rs("singer")%>  <a href="SongAdd.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs("SClassid")%>&NClassid=<%=rs("NClassid")%>&specialid=<%=rs("specialid")%>">[进入添歌]</a></td>
          <td width="9%" align=center><a href="SongSave2.asp?act=SetIsGood&id=<%=rs("id")%>&classid=<%=classid%>&SClassid=<%=SClassid%>&Nclassid=<%=Nclassid%>&page=<%=CurrentPage%>"><%if rs("IsGood")=true then%>撤销<%else%>推荐<%end if%></a></td> 
          <td width="9%" align=center><a href="SongModify.asp?id=<%=rs("id")%>&AskClassid=<%=classid%>&AskNClassid=<%=Nclassid%>&page=<%=CurrentPage%>">修改</a></td> 
          <td width="9%" align=center><a href="SongDel.asp?id=<%=rs("id")%>&classid=<%=classid%>&SClassid=<%=SClassid%>&Nclassid=<%=Nclassid%>&page=<%=CurrentPage%>">删除</a></td> 
        </tr>
<%
	i=i+1
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
                  </table>
<!-----------------------搜索专辑----------------------------->
<%elseif stype="Special" then%>
                   <table border="1" width="90%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
              <tr>
                <td width="33%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">专辑名称</font></td>
                <td width="12%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">所属歌手</font></td>
                <td width="11%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">快速添歌</font></td>
                <td width="11%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">首页推荐</font></td>
                <td width="7%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">修改</font></td>
                <td width="7%" height=22 align=center bgcolor="#FF1171"><font color="#FFFFFF">删除</font></td>
              </tr>
<%
Set Trs= Server.CreateObject("ADODB.Recordset")
do while not rs.eof
	i=i+1
	Tsql="SELECT * FROM MusicList where Specialid="+cstr(rs("Specialid"))
	Trs.open Tsql,conn,1,1
	TotalMNum=Trs.recordcount
	Trs.close
%>
              <tr>
                <td width="33%"><a href="SongList.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs("SClassid")%>&NClassid=<%=rs("NClassid")%>&specialid=<%=rs("specialid")%>"><font color=black><%=rs("name")%></font></a>&nbsp; <a href="SongAdd.asp?Classid=<%=rs("Classid")%>&SClassid=<%=rs("SClassid")%>&NClassid=<%=rs("NClassid")%>&specialid=<%=rs("specialid")%>">[添歌]</a></td>
                <td width="12%" align=center><%=rs("Nclass")%>　</td>
                <td width="11%" align=center><a href="SongConAdd.asp?specialid=<%=rs("specialid")%>">快速添歌</a></td>
                <td width="11%" align=center><a href="AddfileSave.asp?act=SetIsGood&Specialid=<%=rs("Specialid")%>&Askclassid=<%=classid%>&AskSClassid=<%=SClassid%>&AskNclassid=<%=Nclassid%>&page=<%=CurrentPage%>"><%if rs("IsGood")=true then%>撤销<%else%>推荐<%end if%></a></td>
                <td width="7%" align=center><a href="Filemodify.asp?Specialid=<%=rs("Specialid")%>&AskClassid=<%=rs("Classid")%>&AskSClassid=<%=rs("SClassid")%>&AskNClassid=<%=rs("NClassid")%>&page=<%=CurrentPage%>">修改</a></td> 
                <td width="7%" align=center><a href="Filedel.asp?Specialid=<%=rs("Specialid")%>&classid=<%=rs("Classid")%>&AskSClassid=<%=rs("SClassid")%>&Nclassid=<%=rs("NClassid")%>&page=<%=CurrentPage%>">删除</a></td> 
              </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
Set Trs=nothing
%>
                  </table>
<!-----------------------搜索会员----------------------------->
<%elseif stype="User" then%>
<div align="center">
<center>
<table border="0" width="86%" cellspacing="1" cellpadding="1">
 <tr>
<td align=center valign=top>
            <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
              <tr>
                <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">ID</font></td>
                <td width="50%" height=25 align=center bgcolor="#FF1171"><font color="white">用户名</font></td>
                <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">修改</font></td>
                <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">锁定</font></td>
                <td width="13%" height=25 align=center bgcolor="#FF1171"><font color="white">删除</font></td>
              </tr>
<%
i=0
do while not rs.eof
	i=i+1
%>
              <tr>
                <td height="22" align=center><%=rs("id")%>　</td>
                <td align=center><a href="UserModify.asp?id=<%=rs("id")%>"><%=rs("UserName")%></a>　</td>
                <td align=center><input type=button name=edit value=修改 onclick="javascript:window.open('UserModify.asp?id=<%=rs("id")%>','_self','')"></td>
                <td align=center><input type=button name=lock value=<%if rs("lockuser")=false then%>"锁定"<%else%>"开锁"<%end if%> onclick="javascript:window.open('UserSave.asp?id=<%=rs("id")%>&act=lock','_self','')"></td>
                <td align=center><input type=button name=del value=删除 onclick="javascript:window.open('UserDel.asp?id=<%=rs("id")%>','_self','')"></td>
              </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
                  </table></div>

<%
else
end if
%>
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
      <div align="center">
        <center>
        <table border="0" width="100%" cellpadding="2" cellspacing="0">
          <tr><form method=Post action="<%=filename%>?classid=<%=classid%>&Nclassid=<%=Nclassid%>">
            <td bgcolor="#ffffff" align="center"><%=Nclass%>&nbsp; 共找到<font color="#ff0000"><b><%=totalnumber%></b></font>个结果
  &nbsp;分<font color="#ff0000"><%=n%></font>页显示
  &nbsp;当前第<font color="#ff0000"><%=CurrentPage%></font>页
&nbsp;<%if CurrentPage<2 then%>上一页&nbsp;
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&keyword=<%=keyword%>&stype=<%=stype%>">上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  下一页
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&keyword=<%=keyword%>&stype=<%=stype%>">下一页</a>
<%end if%>
  &nbsp;转到:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>第<%=i%>页</option>   
<%next%>   
  </select>       
</td></form>
</tr>
</table>
</center>
</div>       
<%        
end function

function Search
%></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  </center>
</div>
<%end function%><%
set rs=nothing
conn.close
set conn=nothing
%></body></html>