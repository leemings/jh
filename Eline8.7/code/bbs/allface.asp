<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	dim n,m,h,s
	dim y,z,show
	stats="论坛头像列表"
	h=10 '每行显示个数
	s=50 '每页显示个数
	M=Forum_userfaceNum '头象个数
	call nav()
	call head_var(2,0,"","")
	call main()
	call footer()
%>
<%sub main()%>
    <table cellpadding=6 cellspacing=1 class=tableborder1 align=center style="table-layout:fixed;word-break:break-all;">
    <tr>
    <th id=tabletitlelink>
<% if m>s then %>
<p align="center"><% for Y=1 to (m+s-1)\s %>
 [<a href=allface.asp?show=<%=Y%>> 头像列表<%=Y%> </a>]  
<% next 
end if%>
    </th>
    </tr>
<tr>
    <td class=tablebody1>
<table width="100%">
    <tr>
<% 	dim p
	i=0
	p= trim(request("show"))
	if p="" or not isInteger(p) then
	p=1
	end if
	z=s*p

	if z/m>1 then
		z=m
	else
		z=s*p
	end if

	for n=s*p+1-s to z
	i=i+1
%>
      <td align="center" height=50><IMG SRC=<%=Forum_info(11)&Forum_userface(n-1)%>><br><%=Forum_userface(n-1)%></td>
<%
	if i=h then
	response.write "</tr><tr>"
	i=0
	end if 
	next
%>
  </table>
    </td>
    </tr></table>
<%end sub%>