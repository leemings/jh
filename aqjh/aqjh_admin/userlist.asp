<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if session("aqjh_adminok")<>true then response.end
search=cstr(server.HTMLEncode(Request.QueryString("search")))
fs=cstr(server.HTMLEncode(Request.QueryString("fs")))
laren=cstr(server.HTMLEncode(Request.QueryString("laren")))
%>
<html><title>天用户列表</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
</head>
<SCRIPT language=JavaScript>
function go1(){alert("无此类排序!");
}
function go(url){
parent.txt.location.href="userlist.asp?fs="+ url +"&search=<%=search%>";
}
</SCRIPT>
<%
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if laren<>"ok" then
	if fs="" then fs="-id"
	if search<>"" then
		rs.Open "select * from 用户 where (姓名 like '%" & search & "%'or 门派 like '%" & search & "%' or 配偶 like '%" & search & "%' or 情人 like '%" & search & "%' or 信箱 like '%" & search & "%' or lastip like '%" & search & "%' or regip like '%" & search & "%') order by " & fs & " desc", conn, 1,1
	else
		rs.Open "Select * From 用户 Order By " & fs & " DESC", conn, 1,1
	end if
else
	name=cstr(Request.Form("name"))
	zsjf=clng(Request.Form("zsjf"))
	rs.Open "Select * From 用户 where 介绍人='" & name & "' and allvalue>" & zsjf & " Order By 等级 DESC", conn, 1,1
	fs="等级" 
end if
if rs.eof and rs.bof then%>
<div align="center">对不起，暂时没有发现任何<%if laren="ok" then%><font color=blue><%=name%>的拉人</font><%end if%>记录！！ <br>[ <a href=# onclick=history.go(-1)>返回上一级</a> ] 
  <%else
RS.PageSize=20'设置每页记录数,可修改
Dim TotalPages
TotalPages = RS.PageCount
If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if
RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'设置最大记录数
Dim Totalcount
Totalcount =INT(RS.recordcount)
StartPageNum=1
do while StartPageNum+10<=CurPage
StartPageNum=StartPageNum+10
Loop
EndPageNum=StartPageNum+9
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount%>
  <br>
</div>
<table border="1" width="100%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
  <tr> 
    <td colspan="12" align=center bgcolor="#d7ebff">用 户 列 表 (按下面按纽进行排序)<%if laren="ok" then%><b><%=name%>拉人资料查看</b><%end if%></td>
  </tr>
  <tr> 
    <td width="5%" align=center  <%if fs="-id" then%>bgcolor="#66CCCC" <%end if%>>
<input name="B12" onClick=go("-id") type="button" value="ID" class="input1">
    </td>
    <td width="10%" align=center  <%if fs="姓名" then%>bgcolor="#66CCCC" <%end if%>><input name="B122" onClick=go("姓名") type="button" value="姓名" class="input1"></td>
    <td width="4%" align=center ><input name="B132" type="button" onClick=go1() value="性别" class="input1"></td>
    <td width="5%" align=center ><input name="B133" type="button" onClick=go1() value="状态" class="input1"></td>
    <td width="9%" align=center <%if fs="门派" then%>bgcolor="#66CCCC" <%end if%>><input name="B126" type="button" onClick=go("门派") value="门派" class="input1"></td>
    <td width="6%" align=center <%if fs="等级" then%>bgcolor="#66CCCC" <%end if%>><input name="B124" type="button" onClick=go("等级") value="等级" class="input1"></td>
    <td width="4%" align=center <%if fs="grade" then%>bgcolor="#66CCCC" <%end if%>><input name="B125" type="button" onClick=go("grade") value="管理" class="input1"></td>
    <td width="9%" align=center <%if fs="allvalue" then%>bgcolor="#66CCCC" <%end if%>><input name="B1262" type="button" onClick=go("allvalue") value="总积分" class="input1"></td>
    <td width="7%" align=center <%if fs="mvalue" then%>bgcolor="#66CCCC" <%end if%>><input name="B1263" type="button" onClick=go("mvalue") value="月积分" class="input1"></td>
    <td width="18%" align=center <%if fs="lasttime" then%>bgcolor="#66CCCC" <%end if%>><input name="B123" type="button" onClick=go("lasttime") value="最后上线时间" class="input1"></td>
    <td width="18%" align=center <%if fs="lastip" then%>bgcolor="#66CCCC" <%end if%>><input name="B134" type="button" onClick=go("lastip") value="最后上线ip" class="input1">
    </td>
    <td width="5%" align=center ><input name="B1312" type="button" onClick=go1() value="操作" class="input1"></td>
  </tr>
  <%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (I<RS.PageSize)
p=p+1%>
  <tr> 
  <tr> 
    <td width="5%" align=center <%if fs="-id" then%>bgcolor="#66CCCC" <%end if%>><%=rs("id")%></td>
    <td width="10%" align=center <%if fs="姓名" then%>bgcolor="#66CCCC" <%end if%> ><%=rs("姓名")%></td>
    <td width="4%" align=center ><%=rs("性别")%></td>
    <td width="5%" align=center ><%=rs("状态")%></td>
    <td width="9%" align=center <%if fs="门派" then%>bgcolor="#66CCCC" <%end if%>><%=rs("门派")%></td>
    <td width="6%" align=right <%if fs="等级" then%>bgcolor="#66CCCC" <%end if%>><%=rs("等级")%></td>
    <td width="4%" align=right <%if fs="grade" then%>bgcolor="#66CCCC" <%end if%>><%=rs("grade")%></td>
    <td width="9%" align=right <%if fs="allvalue" then%>bgcolor="#66CCCC" <%end if%>><%=rs("allvalue")%></td>
    <td width="7%" align=right <%if fs="mvalue" then%>bgcolor="#66CCCC" <%end if%>><%=rs("mvalue")%></td>
    <td width="18%" align=left <%if fs="lasttime" then%>bgcolor="#66CCCC" <%end if%>><%=rs("lasttime")%></td>
    <td width="18%" align=left <%if fs="lastip" then%>bgcolor="#66CCCC" <%end if%>><%=rs("lastip")%></td>
    <td width="5%" align=center ><a href="SHOWUSER.asp?username=<%=rs("姓名")%>">详细</a></td>
  </tr>

  </tr>
  <%I=I+1
RS.MoveNext
Loop%>
  <tr> 
    <td colspan=14 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff"> 
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr> 
          <form method="GET" action="userlist.asp">
            <td width="50%"> 　关键字： 
              <input type="text" name="search" size=15 class="input1" onfocus=this.select() onmouseover=this.focus()> 
              <input name="B1" type="submit" value="搜索" class="input1">
            </td>
          </form>
          <td width="50%" align="right"> 有<font color=blue><%=Totalcount%></font>用户 共<%=int(Totalcount/25)+1%>页 [ 
            <% if StartPageNum>1 then %>
            <a href="userlist.asp?CurPage=<%=StartPageNum-1%>"><<</a> 
            <%end if%>
            <% For I=StartPageNum to EndPageNum
if I<>CurPage then %>
            <a href="userlist.asp?CurPage=<%=I%>"><%=I%></a> 
            <% else %>
            <b><font color=#ff0000><%=I%></font></b> 
            <% end if %>
            <% Next %>
            <% if EndPageNum<RS.Pagecount then %>
            <a href="userlist.asp?CurPage=<%=EndPageNum+1%>">>></a> 
            <%end if%>
            ] </td>
        </tr>
      </table></td>
  </tr>
</table>
<p align=center>≮快乐江湖总站≯</p></body></html>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>