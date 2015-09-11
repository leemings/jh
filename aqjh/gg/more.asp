<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
%>
<html>
<title>爱情江湖公告</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head><body background=images/bg.gif>
<%
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
	CurPage = 1
Else
	CurPage = CINT(Request.QueryString("CurPage"))
End If
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.Open "Select * From gg Order By e DESC", conn, 1,1
if rs.eof and rs.bof then%>
	暂时没有任何记录！！
<%else
RS.PageSize=25'设置每页记录数,可修改          
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
<table border="1" width="70%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
  <tr> 
    <td bordercolorlight="#00215a" align=center bgcolor="#d7ebff" width="86%"><B>查询公告</B></td>
    <td bordercolorlight="#00215a" align=center bgcolor="#d7ebff" width="14%">
    <%if aqjh_grade=10 then%><a href="add.asp" target=_blank>增加公告</a><%else%><%=date()%><%end if%></td>
  </tr>
  <%I=0
p=RS.PageSize*(Curpage-1)              
do while (Not RS.Eof) and (I<RS.PageSize)              
p=p+1%>
  <tr> 
    <td bordercolorlight="#00215a" align=left colspan="2"><img src=images/046.gif>{<font color=blue><%=rs("d")%>类</font>}<a href="show.asp?id=<%=rs("id")%>" target="_blank"><%=rs("a") %></a><font color=#666666>[<%=rs("e")%>]</font>
<%if aqjh_grade=10 then%><a href="change.asp?act=del&id=<%=rs("id")%>"><font color=red>删除</font></a>|<a href="modifile.asp?id=<%=rs("id")%>"><font color=red>修改</font></a><%end if%>
    </td>
  </tr>
  <%I=I+1              
RS.MoveNext              
Loop%>
  <tr> 
    <td colspan=5 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff"> 
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr> 
          <form method="GET" action="search.asp">
            <td width="50%"> 　关键字： 
              <input type="text" name="search" size=15 class="input1" onfocus=this.select() onmouseover=this.focus()>
              <input name="B1" type="submit" value="搜索" class="input1">
            </td>
          </form>
          <td width="50%" align="center"> 共<%=Totalcount%>条公告 [ 
            <% if StartPageNum>1 then %>
            <a href="more.asp?CurPage=<%=StartPageNum-1%>"><<</a> 
            <%end if%>
            <% For I=StartPageNum to EndPageNum   
		      if I<>CurPage then %>
        		    <a href="more.asp?CurPage=<%=I%>"><%=I%></a> 
            <% else %>
		            <b><font color=#ff0000><%=I%></font></b> 
            <% end if %>
            <% Next %>
            <% if EndPageNum<RS.Pagecount then %>
            <a href="more.asp?CurPage=<%=EndPageNum+1%>">>></a> 
            <%end if%>
            ] </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p align=center>爱情江湖</p></body></html>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%> 