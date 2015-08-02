<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<bgsound src="mid/bg.mid">
<LINK href="style.css" rel=stylesheet>
<script language=javascript>
function check(){
if(document.forms[0].username.value.length<=1){alert('请填入您已经申请成功的用户名称！');document.forms[0].username.select();return false;}
else if(document.forms[0].username.value.indexOf("\'")!=-1||document.forms[0].username.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].username.select();return false;}
else if(document.forms[0].password.value.length<6){alert('密码过短，请验证是否输入正确');document.forms[0].password.select();return false;}
else if(document.forms[0].password.value.indexOf("\'")!=-1||document.forms[0].password.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].password.select();return false;}
else {return true;}
}
</script>
</head>
<body onload="javascript:document.forms[0].username.focus();" bgcolor="<%=bgcolor%>" background="<%=bgimage%>">
<div align=center> 
  <form action="login.asp" method=post onsubmit='return(check());' name=form1>
    <table width="609">
      <tr> 
        <td align="center" valign="middle" width="131"><img src="images/STAMP.gif" width="131" height="280"></td>
        <td align="center" valign="middle" width="450"> 
          <p><img src="images/poem3.gif" width="281" height="170"></p>
          <p><img src="images/partner.gif" width="450" height="60"></p>
        </td>
      </tr>
      <tr>
		<td valign="middle" align=center width="581" colspan="2">姓名<input type="text" name="username" value="" size=14 maxlength=14>
          密码  
          <input type="password" name="password" value="" size=14 maxlength=14> 
          <input type="submit" value=" 登 录 "> 
          <script language=javascript src="onlinelist.asp"></script> 
        </td> 
          
      </tr> 
       <tr ><td width="595" align="center" colspan="2">&nbsp;<input onClick="javascript:window.open('regstep1.asp','register','left=200,top=100,width=500,height=360,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=注册帐号 type=button value="注  册" name="button"> 
          <input onClick="javascript:window.open('relive.asp','relive',' width=300,height=210,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=帐号复活 type=button value="复  活" name="button">     
          <input onClick="javascript:window.open('suicide.asp','suicide',' width=300,height=350,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=自 杀 type=button value="自  杀" name="button">     
          <input onClick="javascript:window.open('zj/help.asp','readme',' width=500,height=300,left=200,top=100,status=no,toolbars=yes,menubars=yes,scrollbars=yes,resize=no')" title=新手入门 type=button value="自 救" name="button">     
          <input onClick="javascript:window.open('faq.htm','faq',' width=500,height=300,left=100,top=100,status=no,toolbars=yes,menubars=yes,scrollbars=yes,resize=no')" title=问题解答 type=button value="问题解答" name="button">     
                    <input onClick="javascript:window.open('chgpw.asp','chgpassword',' width=300,height=300,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=更改密码 type=button value="更改密码" name="button">     
          <input onClick="javascript:window.open('chgdatu1.asp','chgdatum',' width=500,height=300,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=资料更新 type=button value="资料更新" name="button">     
        </td></tr>     
    </table>     
    　 
    <table width="100%" border="0" align="center"> 
       
      <tr> 
        <td align="center" valign="middle">程序修改：<a href="http://www.tl99.com"><font color="#993300">世纪联网</font></a>   
          授权给：<font color="#993300">   
          <% response.write Application("Ba_jxqy_userright")%>   
          </font> 序列号：<font color="#993300">   
          <% response.write Application("Ba_jxqy_seriesnumber")%>   
          </font>开放时间：<font color="#993300">   
          <%=Application("Ba_jxqy_opendata")%></font> 
          访问统计：<font color="#993300">  
          <%=Application("Ba_jxqy_visitor")%>  <Br>
          </font>会员功能：<font color="#993300">  
          <%if Application("Ba_jxqy_fellow")=true then%>
          开放
          <%else%>
          关闭
          <%end if%>
          </font>  
        </td>  
      </tr>  
    </table>  
    <br>  
    <table width="611" border="0">  
      <tr>  
        <td height="29"><font color="#CC0000">【神话虚幻联盟】</font>‘世纪联网’以后不会在更换江湖版本！欢迎各个神话江湖的站长和我联盟共同发展！
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>您的位置</b></font></td>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>您的位置</b></font></td>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>您的位置</b></font></td>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>您的位置</b></font></td>
              <td width="16%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>您的位置</b></font></td>
              <td width="16%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>您的位置</b></font></td>
            </tr>
          </table>
        </td>  
      </tr>  
    </table>  
  </form>  
</div>  
</body>  
