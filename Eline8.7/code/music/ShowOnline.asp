<!--#include file="Config.asp" -->
<%
  const MaxPerPage=45
   	dim totalPut
   	dim CurrentPage
   	dim TotalPages
   	dim i,j
   	dim sql
   	dim rs
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
  set Rs=server.createobject("adodb.recordset")
%>
<HTML><HEAD><TITLE><%= Title_Name %><%= CategoryName_CHS %> ==>> �����û� </TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta name=keywords content="һ������,һ������,���ֽ���,eline_email@etang.com,eline,51eline,happyjh.com">
<LINK href="style.css" type=text/css rel=stylesheet>
</HEAD>
<BODY text=#003300 vLink=#002200 bgColor=#cccccc leftMargin=0 topMargin=0>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td colspan="7" valign="top" bgcolor="#FFFFFF" ><table width="98%" border="0" align="center">
        <tr> 
          <td><img src="images/dotdb.gif" width="10" height="10" align="absmiddle"> 
            ��ǰλ�ã�<a class=white_bg 
            href="/">��ҳ</a> &gt;&gt; <a class=white_bg 
            href="./"><%= CategoryName_CHS %></a> &gt;&gt; �����û�</td>
        </tr>
      </table>
      <table width="98%" border=0 align="center" cellpadding=1 cellspacing=1 bgcolor="#cccccc">
        <tr bgcolor="#E4E4E4"  height=25> 
          <td align="center">�û�</td>
          <td align="center">����ʱ��</td>
          <td align="center"  nowrap>�ʱ��</td>
          <td align="center"  nowrap>�û�IP</td>
          <td align="center"  nowrap>����ϵͳ</td>
          <td align="center"  nowrap>�����</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="6"> <% 
	  sql="select * from "&CategoryName&"_Online order by lastimebk Desc"
       rs.open sql,conn,1,1 
  	  if rs.eof and rs.bof then 
       		response.write "<table><tr><td border=""0"" width=""100%"" height=""100%"" cellspacing=""1"" cellpadding=""0"" bgcolor=""#FFFFFF""><p align=""center"">û�л�û���ҵ��κ������û���</p></td></tr></table>" 
      else 
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
           		showpage totalput,MaxPerPage,"ShowOnline.asp"
            		showContent
            		showpage totalput,MaxPerPage,"ShowOnline.asp"
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           			showpage totalput,MaxPerPage,"ShowOnline.asp"
            			showContent
             			showpage totalput,MaxPerPage,"ShowOnline.asp"
        		else
	        		currentPage=1
           			showpage totalput,MaxPerPage,"ShowOnline.asp"
           			showContent
           			showpage totalput,MaxPerPage,"ShowOnline.asp"
	      		end if
	   	end if
   	rs.close 	
   	end if 
	         
   	sub showContent 
       	dim i 
	   	i=0 
%> </td>
        </tr>
          <% do while not rs.eof %>
          <tr bgcolor="#FFFFFF"> 
            <td align="center"><%= rs("UserName") %></td>
            <td align="center"><%= rs("startime") %></td>
            <td align="center"><%= rs("lastimebk") %></td>
            <td align="center"><%= rs("ip") %></td>
            <td align="center"><%= system(rs("browser")) %></td>
            <td align="center"><%= browser(rs("browser")) %></td>
          </tr>
          <%
	      i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	    loop
	%>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="6"> <%
   end sub 

	function showpage(totalnumber,maxperpage,filename)
  	dim n

  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1
  	end if
  	response.write "<table cellspacing=1 width='100%' border=0 colspan='4' ><form method=Post action="""&filename&"""><tr><td align=right> "
  	if CurrentPage<2 then
    		response.write "��ǰ����<strong><font color=red>"&totalnumber&"</font></strong>λ�û�����&nbsp;��ҳ ��һҳ&nbsp;"
  	else
    		response.write "��ǰ����<strong><font color=red>"&totalnumber&"</font></strong>λ�û�����&nbsp;<a href="&filename&"?page=1&"">��ҳ</a>&nbsp;"
    		response.write "<a href="&filename&"?page="&CurrentPage-1&">��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		response.write "��һҳ βҳ"
  	else
    		response.write "<a href="&filename&"?page="&(CurrentPage+1)&">"
    		response.write "��һҳ</a> <a href="&filename&"?page="&n&">βҳ</a>"
  	end if
   	       response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "
           response.write "&nbsp;<b>"&maxperpage&"</b>λ�û�/ҳ "
%>
            ת���� 
            <select name='page' size='1' style="font-size: 9pt" onChange='javascript:submit()'>
              <%for i = 1 to n%>
              <option value='<%=i%>' <%if CurrentPage=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>
              <%next%>
            </select> <%   
	response.write "</td></tr></FORM></table>"
end function
%> </td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
</BODY></HTML>
<%

  set rs=nothing
CloseDatabase
%>