<!--#include file="conn.asp"-->
<html>
<head>
<title>动画频道</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0" link="#000000" vlink="#000000" alink="#000000">
<!--#include file="top.asp"-->
<table align=center border=0 cellpadding=0 cellspacing=0 width=776>
  <tbody> 
  <tr> 
    <td class=left_bg valign=top width=154 bgcolor="#ffcc00"> 
      <table border=0 cellpadding=0 cellspacing=0 id=NavTab width="100%">
        <tbody> 
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td bgcolor=#e80202 height=38> 
            <div align=center><a href="Index.asp"><img 
            border=0 height=30 src="images/lm.gif" 
width=150></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal id=C10463000010 valign=center> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=1"><img border=0 
            height=23 src="images/l_mu1.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=3"><img border=0 
            height=23 src="images/l_mu2.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=3"><img border=0 
            height=23 src="images/l_mu3.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=4"><img border=0 
            height=23 src="images/l_mu4.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=9"><img border=0 
            height=23 src="images/l_mu5.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=10"><img border=0 
            height=23 src="images/l_mu6.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#e80202 class=left_title height=35> 
            <div align=center>观看排行榜</div>
          </td>
        </tr>
        <tr> 
          <td class=left_line height=5></td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=6 cellspacing=0 class=left_text 
        width="100%">
        <tbody> 
        <tr> 
          <td class=line_h22> 
            <%      
dim rs1,sql1      
set rs1=server.createobject("adodb.recordset")      
sql1="select * from flash ORDER BY hits desc"      
rs1.open sql1,conn,1,1      
if rs1.eof and rs1.bof then       
response.write "<p align='center'>没有热点动画</p>"      
else       
do while not rs1.eof      
dim k      
%>
            <span class="b">·</span><a href="show.asp?ID=<%=rs1("ID")%>&ClassID=<%=rs1("ClassID")%>"><%=rs1("title")%></a> 
            [<span class="s"><font color="#FF0000"><%=rs1("hits")%></font></span>]<br>
            <%k=k+1      
	if k=10 then exit do      
	rs1.movenext      
	loop      
	end if      
	rs1.close      
%>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#e80202 class=left_title height=35> 
            <div align=center>下载排行榜</div>
          </td>
        </tr>
        <tr> 
          <td class=left_line height=5></td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=6 cellspacing=0 class=left_text 
        width="100%">
        <tbody> 
        <tr> 
          <td class=line_h22> 
            <%      
set rs1=server.createobject("adodb.recordset")      
sql1="select * from flash ORDER BY download desc"      
rs1.open sql1,conn,1,1      
if rs1.eof and rs1.bof then       
response.write "<p align='center'>没有动画下载</p>"      
else       
do while not rs1.eof      
dim n      
%>
            <span class="b">·</span><a href="show.asp?ID=<%=rs1("ID")%>&ClassID=<%=rs1("ClassID")%>"><%=rs1("title")%></a> 
            [<span class="s"><font color="#FF0000"><%=rs1("download")%></font></span>]<br>
            <%n=n+1      
	if n=10 then exit do      
	rs1.movenext      
	loop      
	end if      
	rs1.close      
	set rs1=nothing      
%>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td bgcolor=#5e4300 width=6></td>
    <td bgcolor=#ef4204 valign=top width=616 align="center"> 
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td rowspan=2 width=88><img border=0 height=30 
            src="images/list.gif" width=98></td>
          <td bgcolor=#ff9000 height=21 width="100%">　 </td>
        </tr>
        <tr> 
          <td bgcolor=#cc3300 height=9></td>
        </tr>
        </tbody> 
      </table>
      <%   
   const MaxPerPage=5   
   dim totalPut      
   dim CurrentPage   
   dim TotalPages   
   dim i,j   
   if not isempty(request("page")) then   
      currentPage=cint(request("page"))   
   else   
      currentPage=1   
   end if   
   if not isEmpty(request("classid")) then   
	classid=request("classid")   
   else   
	classid=1   
   end if   
   
%>
      <%   
sql="select * from flash where classid="+cstr(classid)+" order by id desc"    
Set rs= Server.CreateObject("ADODB.Recordset")   
rs.open sql,conn,1,1   
  if rs.eof and rs.bof then   
       response.write "<p align='center' class='font'> 还 没 有 任 何 动 画</p>"   
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
            showpage totalput,MaxPerPage,"FlashClass.asp"   
       else   
          if (currentPage-1)*MaxPerPage<totalPut then   
            rs.move  (currentPage-1)*MaxPerPage   
            dim bookmark   
            bookmark=rs.bookmark   
            showContent   
             showpage totalput,MaxPerPage,"FlashClass.asp"   
        else   
	        currentPage=1   
           showContent   
           showpage totalput,MaxPerPage,"FlashClass.asp"   
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
      <table border=0 cellpadding=0 cellspacing=0 width=610>
        <tbody> 
        <tr> 
          <td colspan=5><img alt="" height=33 
                  src="images/test_01.gif" width=610></td>
        </tr>
        <tr> 
          <td rowspan=3><img alt="" height=191 
                  src="images/test_02.gif" width=40></td>
          <td bgcolor=#ffffff rowspan=3 valign=top width=308> 
            <table border=0 cellpadding=8 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td><font color="#666600">动画名称：</font><font color="#FF0000"><%=rs("title")%></font></td>
              </tr>
              <tr> 
                <td><font color="#666600">上传时间：</font><span class=time><%=rs("time")%></span> 
                  <font color="#666600">观看：</font><span class=time><%=rs("hits")%></span> 
                  次 <font color="#666600">评价：</font><%=rs("commend")%></td>
              </tr>
              <tr> 
                <td><font color=#666600>作者信箱：</font> 没有该动画作者的信箱 <font color="#666600">大小：</font><span class=time><%=rs("KB")%>K</span></td>
              </tr>
              <tr> 
                <td><font color="#666600">动画简介：</font> 
                  <% if len (rs("content"))>30 then %>
                  <%= left (rs("content"),28)%>... 
                  <% else %>
                  <%=rs("content")%> 
                  <% end if %>
                </td>
              </tr>
              <tr> 
                <td align="center"><img src="images/img_1.gif" align="absmiddle"> 
                  <a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>">在线观看</a>　　　　<img src="images/download.gif" align="absmiddle"> 
                  <a href="download.asp?ID=<%=rs("ID")%>" target="_blank">下载存盘</a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td rowspan=3><img alt="" height=191 
                  src="images/test_04.gif" width=44></td>
          <td colspan=2><img alt="" height=37 
                  src="images/test_05.gif" width=218></td>
        </tr>
        <tr> 
          <td bgcolor=#ffffff width=188 align="center"><img src="<%=rs("img")%>" width="102" height="102" border="1"></td>
          <td><img alt="" height=116 src="images/test_07.gif" 
                  width=30></td>
        </tr>
        <tr> 
          <td colspan=2><img alt="" height=38 
                  src="images/test_08.gif" width=218></td>
        </tr>
        <tr> 
          <td colspan=5><img alt="" height=32 
                  src="images/test_09.gif" width=610></td>
        </tr>
        </tbody> 
      </table>
      <% i=i+1   
if i>=MaxPerPage then exit do   
rs.movenext   
loop   
%>
      <table width="594" border="0" cellspacing="0" cellpadding="0" align="right">
        <tr> 
          <td height="21" align="center"> 
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
    response.write "<a href="&filename&"?page=1&classid="&classid&"><img src='images/page_top.gif' border=0 align=absmiddle></a>&nbsp;"   
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&classid="&classid&"><img src='images/page_pv.gif' border=0 align=absmiddle></a>&nbsp;"   
  end if   
  if n-currentpage<1 then   
    response.write ""   
  else   
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&classid="&classid&">"   
    response.write "<img src='images/page_next.gif' border=0 align=absmiddle></a> <a href="&filename&"?page="&n&"&classid="&classid&"><img src='images/page_end.gif' border=0 align=absmiddle></a>"   
  end if   
   response.write "&nbsp;页次：</font><b><font color=red>"&CurrentPage&"</font>/"&n&"</b>页</font> "   
    response.write "&nbsp;共<b>"&totalnumber&"</b>个动画 <b>"&maxperpage&"</b>个动画/页"   
end function   
%>
        </tr>
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<!--#include file="bottom.asp"-->   
</body>   
</html>   
