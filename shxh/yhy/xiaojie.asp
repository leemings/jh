<%my=session("Ba_jxqy_username")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>海阔天空江湖-怡红院</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE border=0 cellPadding=0 cellSpacing=0 width=552 align="center">
  <TBODY> 
  <TR>
    <TD background=tdbg1.gif width="11"><IMG src="empty.gif"></TD>
    <TD align=middle vAlign=top width=661>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="601">
        <TBODY>
        <TR>
          <TD width="11"><IMG src="empty.gif" width="1" height="1"></TD>
          <TD noWrap width="557">
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="553">
              <TBODY>
              <TR>
                <TD width="30"><IMG src="icont1.gif"></TD>
                <TD width="519">
                  <TABLE border=0 cellPadding=0 cellSpacing=0 height="26" width="518">
                    <TBODY>
                    <TR>
                      <TD rowSpan=3 height="26" width="14"><IMG src="empty.gif"></TD>
                      <TD rowSpan=3 height="26" width="11"><IMG src="wtdr1.gif"></TD>
                      <TD background=../pic/wtdrbg1.gif height="5" width="144">
                      <IMG 
                        src="empty.gif"></TD>
                      <TD rowSpan=3 height="26" width="9"><IMG src="wtdr2.gif"></TD>
                      <TD rowSpan=3 height="26" width="330" align="right">
                      <IMG src="empty.gif"><a href="javascript:window.close();"><img border="0" src="close.gif"></a></TD></TR>
                    <TR>
                      <TD align=middle class=redtext height=15 width="144">::<font size="2" color="#FD794D"><b>怡红院</b></font>::</TD>
                    </TR>
                    <TR>
                      <TD background=wtdrbg2.gif height="6" width="144">
                      <IMG 
                        src="empty.gif"></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
          <TD vAlign=bottom width="28"><IMG src="rightct12.gif"></TD>
          <TD vAlign=bottom width="19"><IMG src="rightct13.gif"></TD></TR>
        <TR>
          <TD width="11"><IMG src="rightct3.gif"></TD>
          <TD background=rightct4.gif width="557">
          <IMG 
            src="empty.gif"></TD>
          <TD background=rightct4.gif width="28">
          <IMG 
            src="empty.gif"></TD>
          <TD width="19"><IMG src="rightct14.gif"></TD></TR>
        <TR>
          <TD background=rightct6.gif rowSpan=2 width="11">
          <IMG 
            src="empty.gif"></TD>
          <TD class=bg1 colSpan=2 rowSpan=2 
          style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" vAlign=top 
            width=557> 
            <table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
              <tr bgcolor="#FFFFFF"> 
                <td height="16" colspan="5"> 
                  <p align=center class="p9"><font color="#FF3333"><b>九天江湖怡红院</b></font></p>
                
              <tr bgcolor=#74E76D bordercolor="#000000"> 
                <td height="15" bgcolor="#FFFFCC" bordercolor="#FF6600" width="11%"> 
                  <div align="center"><font size="2" color="#FF0000">名妓</font></div>
                </td>
                <td height="15" bgcolor="#FFFFCC" bordercolor="#FF6600" width="12%"> 
                  <div align="center"><font size="2" color="#FF0000">美貌</font></div>
                </td>
                <td height="15" width="20%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
                  <div align="center"><font size="2" color="#FF0000">和其聊天的价格</font></div>
                </td>
                <td height="15" width="32%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
                  <div align="center"><font size="2" color="#FF0000">和其聊天增加的体力</font></div>
                </td>
                <td height="15" width="25%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
                  <div align="center"><font size="2" color="#FF0000">超做</font></div>
                </td>
              </tr>
              <!--#include file="jiu.asp"-->
              <%
sql="SELECT * FROM 名妓"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
              <tr bgcolor=#DEAD63> 
                <td bgcolor="#FFFFCC" width="11%"> 
                  <center>
                    <font size="2"><%=rs("姓名")%><span class="p9"><font color="#6666FF"></font></span> 
                    </font> 
                  </center>
                  <font size="2"></font> 
                <td bgcolor="#FFFFCC" width="12%"> 
                  <div align="center"><font size="2"><%=rs("美貌度")%></font> </div>
                
                <td bgcolor="#FFFFCC" width="20%"> 
                  <div align="center"><font size="2"><%=rs("美貌度")*1.5%> </font></div>
                </td>
                <td bgcolor="#FFFFCC" width="32%"> 
                  <div align="center"><font size="2"><%=rs("美貌度")/2%> </font></div>
                </td>
                <td bgcolor="#FFFFCC" width="25%"> 
                  <center>
                    <font size="2"><a href=girl.asp?id=<%=rs("id")%>><span class="calen-curr">进入其闺房</span></a></font> 
                  </center>
                </td>
              </tr>
              <%
rs.movenext
loop
set rs=nothing
connt.close%>
            </table>
          </TD>      
          <TD background=rightct8.gif vAlign=top width="19">
          <IMG       
            src="rightct15.gif"></TD></TR>      
        <TR>      
          <TD background=../pic/rightct8.gif width="19">
          <IMG       
            src="empty.gif"></TD></TR>      
        <TR>      
          <TD width="11"><IMG src="rightct9.gif"></TD>      
          <TD background=rightct10.gif colSpan=2 width="587">
          <IMG       
            src="empty.gif"></TD>      
          <TD width="19"><IMG src="rightct11.gif"></TD></TR></TBODY></TABLE>      
    </TD>     
    <TD background=tdbg2.gif width="13"><IMG src="empty.gif"></TD></TR>     
  <TR>     
    <TD width="11"><IMG src="tdbgr1.gif"></TD>     
    <TD background=tdbg3.gif width="661">　</TD>     
    <TD width="13"><IMG src="tdbgr2.gif"></TD></TR></TBODY></TABLE>     
<p align="center"><a target="_blank" href="http://www.junefield.net/bbs/">
海阔天空综合网</a></p>
</BODY></HTML>