<!--#include file="conn.asp"-->
<html>
<title>E线江湖总站___DJ舞吧___wWw.51eline.com</title>
<body style='background:transparent'>
<script src="js/js.js"></script>
<style>
td
{
FONT-WEIGHT: normal; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; color: #000000
}
A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #000000
}
A:visited {
	COLOR: #333333
}
A:active {
	COLOR: #FF0000
}
A:hover {
	COLOR: #333333; TEXT-DECORATION: underline overline
}


</style>

<%
if (isnull(session("MusicUser")) and isnull(session("MusicPwd")) and session("MusicUser")="" and session("MusicPwd")="") and Session("IsAdmin")=true and session("KEY")="super" then
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('为避免给注册用会带来不必要的骚扰，所以会员信息仅对本站注册用户开放!\n如果您已经注册，请先登陆然后再查看，如果您没有注册，请先注册!');window.close();</script>"
	Response.End 
end if
if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
%>
<div align="center">
  <center>
  <table border="1" cellspacing="0" bordercolor="#666666" width="100%" id="AutoNumber1" height="1" style="border-collapse: collapse" cellpadding="0" bgcolor="#FFCC66">
    <tr>
      <td width="755" height="1" bgcolor="#EC9F00">
      <p align="center">-==本站会员列表==-</td>
    </tr>
    <tr>
      <td width="100%" height="1">
      <div align="center">
        <center>
     <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td width="70%" valign="top" align="center">


　</td>
          </tr>
        </form>



</td>
  </tr>
  <tr>
    <td width="70%" valign="top" align="center">
<%
set rs=server.createobject("adodb.recordset")
sql="select * from user where username<>'短信精灵' order by id desc" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
	response.write "<p align='center'>暂时没有用户注册</p>" 
else 
	MaxPerPage=18
	PageUrl="UserList.asp"
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
            <table border="1" width="100%" cellspacing="0" bordercolor="#666666" height="1" cellpadding="0" style="border-collapse: collapse">
              <tr>
                <td width="28" height=1 align=center bgcolor="#EC9F00">
                ID</td>
                <td width="94" height=1 align=center bgcolor="#EC9F00">
                用户名</td>
                <td width="36" height=1 align=center bgcolor="#EC9F00">
                性别</td>
                <td width="155" height=1 align=center bgcolor="#EC9F00">
                电子邮箱</td>
                <td width="72" height=1 align=center bgcolor="#EC9F00">
                OICQ</td>
                <td width="68" height=1 align=center bgcolor="#EC9F00">
                主页</td>
                <td width="136" height=1 align=center bgcolor="#EC9F00">
                发短消息</td>
              </tr>
<%
do while not rs.eof
	i=i+1
	id=rs("id")
%>
              <tr>
                <td align=center width="28" height="1"><%=id%></td>
                <td align=center width="94" height="1"><a href="javascript:open_window('UserInfo.asp?id=<%=rs("id")%>','UserInfo','scrollbars=no,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,copyhistory=no,top=10,left=10,width=300,height=200')"><%=rs("UserName")%></a><%if DateValue(rs("regDate"))=date() then%> <img src="images/new.gif"><%end if%>
<%if Session("IsAdmin")=true and session("KEY")="super" then%>
<br>
<a href="cnadmin/UserModify.asp?id=<%=rs("id")%>"><font color="#FF0000">修</font></a><font color="#FF0000">
                </font>
<a href="cnadmin/UserDel.asp?id=<%=rs("id")%>"><font color="#FF0000">删</font></a><font color="#FF0000">
                </font>
<%end if%>
                </td>
                <td align=center width="36" height="1"><%if rs("sex")=true then%>男<%else%>男<%end if%></td>
                <td align=center width="155" height="1"><%=rs("Email")%></td>
                <td align=center width="72" height="1"><%if rs("oicq")<>"" then%><%=rs("oicq")%><%else%>无<%end if%></td>
                <td align=center width="68" height="1"><%if rs("homepage")<>"" then%><a target="_blank" href="http://<%=rs("homepage")%>">点此查看</a><%else%>无<%end if%></td>
                <td align=center width="136" height="1">
                <a href="javascript:open_window('messager.asp?action=new&touser=<%=rs("UserName")%>','messanger','width=420,height=290')"><font color="#FF0000">
          给</font> 
            <%=rs("UserName")%> <font color="#FF0000">发短消息</font></a></td>
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
<form method=Post action="<%=filename%>">
  <center>
  <p>共<font color="red"><b><%=totalnumber%></b></font>位用户
<%if CurrentPage<2 then%> &nbsp;首页 上一页&nbsp;
<%else%> &nbsp;<a href="<%=filename%>?page=1">首页</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>">
  上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%> 下一页 末页
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>">
  下一页</a>
  <a href="<%=filename%>?page=<%=n%>">
  末页</a>
<%end if%> &nbsp;页次:<strong><font color="red"><%=CurrentPage%>/<%=n%></font></strong>页 
  转到:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>
  第<%=i%>页</option>   
<%next%>   
  </select> </p>
</form>        
<%end function%>
    </td>
  </tr>
  </table></center>
      </div>
      </td>
    </tr>
    <tr>
      <td width="750" height="7" bgcolor="#EC9F00">
      &nbsp;
      </td>
    </tr>
  </table>
  </center>
</div>
<%
set rs=nothing
conn.close
set conn=nothing%></body></html>