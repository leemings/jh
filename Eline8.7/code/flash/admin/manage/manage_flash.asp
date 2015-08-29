<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<title>管理动画</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<link rel="stylesheet" href="../style.css" type="text/css">
<script>
function del(id){
if(confirm('您确定要删除吗？'))
 {
 location.href="del_flash.asp?id="+id+""
 }
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000"  background="bkg.gif">
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#F7F7F7" align="center"><%
   const MaxPerPage=18
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j
   if not isempty(request("page")) then
      currentPage=cint(request("page"))
   else
      currentPage=1
   end if
%>
<%
sql="select * from Flash order by id desc" 
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
  if rs.eof and rs.bof then
       response.write "<p align=center class=font>暂无动画</p>"
   else
	  totalPut=rs.recordcount
      totalPut=rs.recordcount
      if currentpage<1 then
          currentpage=1
      end if
      if (currentpage-1)*MaxPerPage>totalput then
	   if (totalPut mod MaxPerPage)=0 then
	     currentpage= totalPut \ MaxPerPage
	   else
	      currentpage= totalPut \ MaxPerPage + 1
	   end if
      end if
       if currentPage=1 then

            showContent
            showpage totalput,MaxPerPage,"manage_flash.asp"
       else
          if (currentPage-1)*MaxPerPage<totalPut then
            rs.move  (currentPage-1)*MaxPerPage
            dim bookmark
            bookmark=rs.bookmark
            showContent
             showpage totalput,MaxPerPage,"manage_flash.asp"
        else
	        currentPage=1
           showContent
           showpage totalput,MaxPerPage,"manage_flash.asp"
	      end if
	   end if
   rs.close
   end if
	        
   set rs=nothing  
   conn.close
   set conn=nothing
   sub showContent
       dim i
	   i=0
%>
      <%do while not rs.eof%>
<tr bgcolor="#FFFFFF">
    <td><%=rs("id")%>.<%=rs("title")%></a></td>
    <td width="60" align="center" bgcolor="#F7F7F7"><a href="edit_flash.asp?id=<%=rs("id")%>&action=edit">编辑</a></td>
    <td width="60" align="center" bgcolor="#F7F7F7"><a href="javascript:del(<%=rs("id")%>)">删除</a></td>
  </tr>
<% i=i+1
if i>=MaxPerPage then exit do
rs.movenext
loop
%>
</table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td height="21" bgcolor="#F7F7F7" align="center">
<%
end sub 
function showpage(totalnumber,maxperpage,filename)
  dim n
  if totalnumber mod maxperpage=0 then
     n= totalnumber \ maxperpage
  else
     n= totalnumber \ maxperpage+1
  end if
  if CurrentPage<2 then
    response.write ""
  else
    response.write "<a href="&filename&"?page=1><img src='../../images/page_top.gif' border=0 align=absmiddle></a>&nbsp;"
    response.write "<a href="&filename&"?page="&CurrentPage-1&"><img src='../../images/page_pv.gif' border=0 align=absmiddle></a>&nbsp;"
  end if
  if n-currentpage<1 then
    response.write ""
  else
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&">"
    response.write "<img src='../../images/page_next.gif' border=0 align=absmiddle></a> <a href="&filename&"?page="&n&"><img src='../../images/page_end.gif' border=0 align=absmiddle></a>"
  end if
   response.write "&nbsp;页次：</font><b><font color=red>"&CurrentPage&"</font>/"&n&"</b>页</font> "
    response.write "&nbsp;共<b>"&totalnumber&"</b>个动画 <b>"&maxperpage&"</b>个动画/页"
end function
%>
    </td>
  </tr>
</table>
</body>
</html>
