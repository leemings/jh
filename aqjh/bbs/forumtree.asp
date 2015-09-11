<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))
sql="Select * From bbsconfig where followid="&id&" and hide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
alltree=""&alltree&"<div class=menuitems><a href=ShowForum.asp?forumid="&rs1("id")&">"&rs1("bbsname")&"</a></div>"
rs1.movenext
loop
Set Rs1 = Nothing
%>
<SCRIPT>
parent.temp<%=id%>.innerHTML="<a onmouseover=\"showmenu(event,'<%=alltree%>')\" href=ShowForum.asp?forumid=<%=id%>><%=Conn.Execute("Select bbsname From [bbsconfig] where id="&id&"")(0)%></a>"
</SCRIPT>
<%responseend%>