<%
randomize timer
s=1+int(rnd*10000)%><html>
<head>
<title><%=Application("yx8_mhjh_systemname")%></title>
<link rel="stylesheet" href="../style.css" type="text/css">
</head>
<body bgcolor="#804040" background="../image/bg3.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="getpw2.asp" method="post">
  <input type=hidden name=reg1 value="<%=s%>">
  <table width="183" cellpadding="3" cellspacing="2" align="center" border="0" class="p9">
    <tr align="center" valign="middle"> 
      <td colspan="2" height="30" width="171"><b>取 回 密 码</b></td>
    </tr>
    <tr align="center" valign="middle"> 
      <td width="37">帐号</td>
      <td width="124" height="40" align="center"> 
        <input name="account" maxlength="14" size="14" style="text-decoration: blink; border: 1px solid #804040" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'">
      </td>
    </tr>
    <tr align="center" valign="middle"> 
      <td width="37">姓名</td>
      <td width="124" height="40" align="center"> 
        <input name="username" maxlength="14" size="14" style="text-decoration: blink; border: 1px solid #804040" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'">
      </td>
    </tr>
                <tr>   
              <td height="25" width="37">&nbsp;认证<font color="#FF0080"> &nbsp;</font></td>     
              <td height="25" width="124" align="center">    
                <input type=text name=reg size=14 maxlength="10" style="text-decoration: blink; border: 1px solid #804040" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'">      
                <br>输入认证码：<%=s%></td>      
            </tr>      
       <tr align="center" valign="middle">     
      <td colspan="2" height="40" width="171">     
        <input type="submit" value=" 拿 回 " name="submit">    
        <input type="button" onClick="javascript:top.window.close();" value=" 关 闭 " name="button">    
      </td>    
    </tr>    
  </table>    
</form>    
</body></html>    
