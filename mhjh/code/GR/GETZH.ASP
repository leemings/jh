<%
randomize timer
s=1+int(rnd*10000)%>
<HTML><HEAD><TITLE>查找帐号</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="reg/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
</HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=30>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=596>
  <TBODY> 
  <TR>
    <TD background=reg/history_table_bg.gif height=158 vAlign=top width="596"> 
      <div align="center">
        <center> 
      <TABLE border=10 cellPadding=0 cellSpacing=0 class=mountain 
      width=596 height="153" bordercolor="#840000" style="border-collapse: collapse; background-image:url('reg/page_bottom_mountain.gif')">
        <TBODY> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="19">
          <table align="center" border="0" cellPadding="0" cellSpacing="0" width="596">
            <tr>
              <td background="reg/history_top_bg.gif" vAlign="top" width="596">
              <table align="center" border="0" cellPadding="2" cellSpacing="0" class="p9" width="97%" height="24">
                <tr>
                  <td height="20" width="31%">　</td>
                  <td height="20" vAlign="top" width="37%">
                  <div align="center">
                    <font size="3" color="#FF0000">查找帐号</font></div>
                  </td>
                  <td height="20" vAlign="top" width="32%">　</td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </TD></TR> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="30" bordercolor="#000000">
          <p align="center"><img border="0" src="reg/top1.gif"></TD></TR> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="45" bordercolor="#000000">
            <div align="center">
              <center>
            <TABLE background=reg/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=p9 width="537" height="195" style="border-collapse: collapse">
              <TBODY> 
    
    </tr>
<form action="getzh2.asp" method="post">
  <input type=hidden name=out1 value="<%=s%>">
    <tr> 
            <td align="right" width="213" height="46">姓名：</td>
      <td align="left" width="324" height="46"> 
        <input name="username" maxlength="14" size="14" style="text-decoration: blink; border: 1px solid #804040" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'">
      </td>
    </tr>
    <tr> 
      <td align="right" width="213" height="34" >密码：</td>
      <td align="left" width="324" height="34" > 
        <input type=password name=password maxlength="14" size="14" style="text-decoration: blink; border: 1px solid #804040" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'">

      </td>
    </tr>
                <tr>   
              <td align="right" width="213" height="67">认证：</td>      
              <td align="left" width="324" height="67">     
                <input type=text name=out size=14 maxlength="10" style="text-decoration: blink; border: 1px solid #804040" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'">       
                <br>输入认证码：<%=s%></td>       
            </tr>       
       <tr>      
      <td colspan="2" width="537" height="48" >      
        <p>      
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      
        <input type="submit" value=" 拿 回 " name="submit">     
        <input type="button" onClick="javascript:top.window.close();" value=" 关 闭 " name="button">     
      </td>     
    </tr>
    <tr> 
</td>
    </tr>
    <form action="relive2.asp" method="post">
              </TBODY></TABLE>
              </center>
          </div>
            </TD></TR> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="44">
          <p align="center"><img border="0" src="reg/top3.gif"></TD></TR></TBODY></TABLE></center>
      </div>
    </TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="596"><IMG height=1 src="reg/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>