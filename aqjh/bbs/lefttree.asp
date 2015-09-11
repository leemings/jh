<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))
if Request("menu")="menu" then

sql="Select * From menu where followid="&id&" order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof

if ii>1 then 
jjj="</tr><tr>"
ii=1
else
jjj=""
ii=ii+1
end if


ClubTreeList=ClubTreeList&""&jjj&"<td height=20><a target=main href="&rs1("url")&">"&rs1("name")&"</a></td>"


rs1.movenext
loop



%>
<SCRIPT>
parent.menu<%=id%>.innerHTML="<table cellspacing=0 cellpadding=0 width=130><tr><%=ClubTreeList%></tr></table>";
</SCRIPT>
<%
responseend
end if




dim alltree
sub TreeList(selec)
sql="Select * From bbsconfig where followid="&selec&" and hide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof

bbsname=rs1("bbsname")
if Len(""&bbsname&"")>10 then
bbsname=left(""&bbsname&"",8)&"..."
end if

if selec=0 then
alltree=""&alltree&"</DIV>¡õ.<A href=javascript:expands('yuzi"&rs1("id")&"')><font color=215DC6>"&bbsname&"</font></A><BR><DIV class=child id=yuzi"&rs1("id")&"Child style='display:none'>"
else
alltree=""&alltree&"¡¡©».<A href=ShowForum.asp?forumid="&rs1("id")&" target=main>"&bbsname&"</A><br>"
end if

ii=ii+1
if ii <2 then TreeList rs1("id")
ii=ii-1

rs1.movenext
loop
Set Rs1 = Nothing
end sub


TreeList(0)

%>
<SCRIPT>
parent.followTd.innerHTML="<%=alltree%></DIV>";
</SCRIPT>

<%responseend%>