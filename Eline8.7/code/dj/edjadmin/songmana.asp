<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<%
if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
sclassid=request.QueryString("Sclassid")
%>
<meta http-equiv="Content-Language" content="zh-cn">
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
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse">
  <tr>
    <td valign=top width=150>
    <!--#include file="admin_left.asp"-->
    　</td>
    <td valign=top width=10>　</td>
    <td valign=top width="600">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3">
      <tr>
        <td width="100%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber4" bgcolor="#FF9900" height="98">
          <tr>
            <td width="100%" height="98" valign="top">
            <div align="center">
              <center>
              <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
                <tr>
                  <td width="100%" align="center" bgcolor="#A65300" colspan="6">
<%
'-----------------目录列表-----------------------
set Trs=server.createobject("adodb.recordset")
Tsql="select SClassid,SClass from SClass"
Trs.open Tsql,conn,1,1
if not Trs.eof then
do while not Trs.eof
%>
<a href="SongMana.asp?SClassid=<%=Trs("SClassid")%>"><%=Trs("SClass")%></a>|
<%
	Trs.movenext
	loop
	end if
	Trs.close
	Set Trs=nothing

%>
<%
set rs=server.createobject("adodb.recordset")
if request.QueryString("Sclassid")="" then
sql="SELECT * FROM musicdj order by id desc"
else
sql="SELECT * FROM musicdj where sclassid="&cstr(Sclassid)
end if
rs.open sql,conn,1,1
if rs.EOF then
	response.write "</td></tr><tr><td>未收录舞曲</td></tr>"
else
	MaxPerPage=20
	PageUrl="songmana.asp"
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
'		showpage totalput,MaxPerPage,PageUrl
		showContent 
		showpage totalput,MaxPerPage,PageUrl
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move  (currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark 
'			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		else 
			currentPage=1 
'			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		end if 
	end if 
end if 
rs.close 
			
sub showContent 
%>

</td>
                </tr>
                <tr>
                  <td width="100%" align="center" bgcolor="#FFECD9" colspan="6">
                  <font color="#111111">舞曲管理</font></td>
                </tr>
                <tr>
                  <td width="39%" bgcolor="#A65300">舞曲名称</td>
                  <td width="18%" bgcolor="#A65300">制作人</td>
                  <td width="22%" bgcolor="#A65300">添加时间</td>
                  <td width="7%" bgcolor="#A65300" align="center">置顶</td>
                  <td width="7%" bgcolor="#A65300">修改</td>
                  <td width="7%" bgcolor="#A65300">删除</td>
                </tr>
<%
i=0
do while not rs.eof
	i=i+1
%>
                <tr>
                  <td width="39%" bgcolor="#FFCE9D"><%=rs("musicname")%></td>
                  <td width="18%" bgcolor="#FFCE9D">DJ Anjo</td>
                  <td width="22%" bgcolor="#FFCE9D"><%=rs("dateandtime")%></td>
                  <td width="7%" bgcolor="#FFCE9D" align="center">
                  <a href="SongSave.asp?act=istop&id=<%=rs("id")%>&Sclassid=<%=Sclassid%>&page=<%=CurrentPage%>"><%if rs("Istop")="1" then%><font color="#FF00FF">撤销</font><%else%><font color="#FF00FF">置顶</font><%end if%></a>
                  </td>
                  <td width="7%" bgcolor="#FFCE9D">
                  <a href="SongEdit.asp?id=<%=rs("id")%>&AskSClassid=<%if Sclassid<>"" then%><%=Sclassid%><%else%><%=rs("SClassid")%><%end if%>&page=<%=CurrentPage%>">
                  <font color="#800000">修改</font></a></td>
                  <td width="7%" bgcolor="#FFCE9D">
                  <a href="Songsave.asp?act=del&id=<%=rs("id")%>&Sclassid=<%=Sclassid%>&page=<%=CurrentPage%>">
                  <font color="#800000">删除</font></a></td>
                </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>

                <tr>
                  <td width="100%" bgcolor="#FFECD9" colspan="6">　</td>
                </tr>
                <tr>
                  <td width="100%" colspan="6">
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<table style="border-collapse: collapse" cellpadding="0" cellspacing="0" border="0" bgcolor="#FFCE9D" width="100%">
<form method=Post action="<%=filename%>?SClassid=<%=SClassid%>">
<tr>
<td align="center">
共<b><%=totalnumber%></b>首舞曲
<%if CurrentPage<2 then%> &nbsp;首页 上一页&nbsp;
<%else%> &nbsp;<a href="<%=filename%>?page=1&SClassid=<%=SClassid%>">首页</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&SClassid=<%=SClassid%>">
  上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%> 下一页 末页
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&SClassid=<%=SClassid%>">
  下一页</a>
  <a href="<%=filename%>?page=<%=n%>&SClassid=<%=SClassid%>">
  末页</a>
<%end if%> &nbsp;页次:<strong><%=CurrentPage%>/<%=n%></strong>页 
  转到:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>
  第<%=i%>页</option>   
<%next%>   
  </select>
</td>
</tr>
</form>
</table>
<%end function%>
</td>
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
    </td>
  </tr>
  </table>
  </center>
</div>