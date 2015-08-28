<%
randomize timer
s=1+int(rnd*10000)
%>

<HTML><HEAD><TITLE>掉线自救</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="reg/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 5.0" name=GENERATOR><script language=javascript>
function check(){
if(document.forms[0].username.value.length<=1){alert('请填入您已经申请成功的用户名称！');document.forms[0].username.select();return false;}
else if(document.forms[0].username.value.indexOf("\'")!=-1||document.forms[0].username.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].username.select();return false;}
else if(document.forms[0].password.value.length<6){alert('密码过短，请验证是否输入正确');document.forms[0].password.select();return false;}
else if(document.forms[0].password.value.indexOf("\'")!=-1||document.forms[0].password.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].password.select();return false;}
else {return true;}
}
</script>
</HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0>
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
                    <font color="#FF0000" size="3">掉线自救</font></div>
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
            cellPadding=0 cellSpacing=0 class=p9 width="537" height="321" style="border-collapse: collapse">
              <TBODY> 
              <TR>
                <TD align=left vAlign=top height="21" width="537"> 
                  <class=p9><font color="#FF0000"><b>          
                  <blockquote>
                    <blockquote>
                      <blockquote>
 <form action="help1.asp" method=post onsubmit='return(check());' name=form1>
<input type=hidden name=reg1 value="<%=s%>">

                       <p><b>自　行　解　挂</b></p>
                      </blockquote>
          <p>姓 名 
            <input type="text" name="username" value="" size=14 maxlength=14>
          </p>
          <p>密 码 
            <input type="password" name="password" value="" size=14 maxlength=14>
          </p>
          <p>认 证 
            <input type=text name=reg size=14 maxlength="14" ><br>
            请输入:<%=s%></p>
          <p> 
            &nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="submit" value="掉线处理">
          </p>
          <p align="left">如果从这里无法成功解挂，请下载<a href="Cookies.bat"><font color="#FFFF00">Cookie.bat</font></a>，运行后再登陆！</p>
                    </blockquote>
                  </blockquote>
         </b></font>
                  </TD></p></TR></TBODY></TABLE>
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