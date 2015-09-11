<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Dim SplitReflashPage
Dim DoReflashPage
dim shuaxin_time
DoReflashPage=Application("tuziji_DoReflashPage")
shuaxin_time=10
ReflashTime=Now()
if (not isnull(session("ReflashTime"))) and cint(shuaxin_time)>0 and DoReflashPage then
	if DateDiff("s",session("ReflashTime"),Now())<cint(shuaxin_time) then
   	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>本页面起用了防刷新机制，请不要在<b><font color=ff0000>"&shuaxin_time&"</font></b>秒内连续刷新本页面<BR>正在打开页面，请稍候……"
	response.end
	else
	session("ReflashTime")=Now()
	end if
elseif isnull(session("ReflashTime")) and cint(shuaxin_time)>0 and DoReflashPage then
	Session("ReflashTime")=Now()
end if
randomize timer
regjm=int(rnd*8998)+1000
%>
<head>
<title><%=Application("aqjh_chatroomname")%></title>
<LINK href="style0.css" rel=stylesheet>
<script language=javascript>

function getOs() 
{ 
    var OsObject = ""; 
   if(isChrome=navigator.userAgent.indexOf("Chrome")!=-1){ 
        return "<audio src='mid/b11g.mp3' type='audio/mp3' autoplay='autoplay' hidden='true'></audio> "; 
   }
   else if(isIE = navigator.userAgent.indexOf("MSIE")!=-1) { 
//        return "<embed autostart='true' loop='-1' controls='ControlPanel' width='0' height='0' src='/mid/bg.mp3'/>"; 
		return "<bgsound src='mid/b11g.mp3'>";
   }
   else {
	   return "<embed autostart='true' loop='-1' controls='ControlPanel' width='0' height='0' src='mid/b11g.mp3'/>"; 
   }
   /*if(isFirefox=navigator.userAgent.indexOf("Firefox")!=-1){ 
        return "Firefox"; 
   } 
   if(isChrome=navigator.userAgent.indexOf("Chrome")!=-1){ 
        return "<audio src='/mid/anqi.wav' type='audio/mp3' autoplay='autoplay' hidden='true'></audio> "; 
   } 
   if(isSafari=navigator.userAgent.indexOf("Safari")!=-1) { 
        return "Safari"; 
   }  
   if(isOpera=navigator.userAgent.indexOf("Opera")!=-1){ 
        return "Opera"; 
   }
   */
} 

document.write(getOs());

function check(){
if(document.forms[0].username.value.length<=1){alert('请填入您已经申请成功的用户名称！');document.forms[0].username.select();return false;}
else if(document.forms[0].username.value.indexOf("\'")!=-1||document.forms[0].username.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].username.select();return false;}
else if(document.forms[0].password.value.length<6){alert('密码过短，请验证是否输入正确');document.forms[0].password.select();return false;}
else if(document.forms[0].password.value.indexOf("\'")!=-1||document.forms[0].password.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].password.select();return false;}
else {return true;}
}
</script>
</head>
<body onload="javascript:document.forms[0].username.focus();" bgcolor="<%=bgcolor%>" background="images/bg1.gif">
<div align=center> 
  <form action="check.asp" method=post onsubmit='return(check());' name=form1>
    <table width="609" border=0>
      <tr> 
        <td align="center" valign="middle" width="131"><img src="images/STAMP.png" width="220" height="350"></td>
        <td align="center" valign="middle" width="450"> 
          <p><img src="images/poem3.gif" width="281" height="170"></p>
          <p><img src="images/partner.gif" width="450" height="60"></p>
        </td>
      </tr>
      <tr>
		<td valign="middle" align=center width="581" colspan="2">姓名<input type="text" name="name" value="" size=14 maxlength=14>
          密码  
          <input type="password" name="pass" value="" size=14 maxlength=14> 
          <input type="submit" value=" 登 录 "> 
		  <script language=JavaScript src="online.asp"></script>
        </td> 
          
      </tr> 
       <tr ><td width="595" align="center" colspan="2">&nbsp;
		  <input onClick="window.open('yamen/read.htm','','width=400,height=502,menubar=no,top=100,left=200')" title=注册帐号 type=button value="注  册" name="button"> 
          <input onClick="window.open('yamen/disp.asp','casper','width=400,height=300,top=200,left=300')" title=帐号复活 type=button value="复  活" name="button">     
		  <input onClick="window.open('yamen/baoshi.asp','baoshi','width=400,height=320,top=200,left=300')" title=保 释 type=button value="保  释" name="button">
          <input onClick="window.open('yamen/getid.asp','','width=400,height=250,menubar=no,top=200,left=300')" title=取回ID type=button value="取回ID" name="button">
          <input onClick="window.open('yamen/close.asp','close','width=400,height=320,top=200,left=300')" title=掉线自救 type=button value="掉线自救" name="button">
          <input onClick="window.open('yamen/modify.asp','modify','width=400,height=350,top=200,left=300')" title=更改密码 type=button value="更改密码" name="button">     
          <input onClick="window.open('yamen/jiejiu.asp','jiejiu','width=400,height=350,top=200,left=300')" title=睡眠自救 type=button value="睡眠自救" name="button">
		  <input onClick="window.open('yamen/gmhx.asp','gmhx','width=400,height=350,top=200,left=300')" title=改名 type=button value="改名" name="button">
		  <input onClick="window.open('yamen/getpass.asp','getpass','width=400,height=350,top=200,left=300')" title=忘记密码 type=button value="忘记密码" name="button">
        </td></tr>     
    </table>     
    　 

    <table width="500" border="0">  
      <tr>  
        <td height="29"><font color="#CC0000">【文字江湖】</font>‘快乐江湖’:回忆匆匆那几年，你是否还记得……
        </td>  
      </tr>  
	  <tr>
	  <td height="29"><font color="#CC0000">本江湖初开，招收副站长，官府管理员，有意者联系站长QQ865240608。另外拉人有奖励哦</font>
	  </tr>
    </table>  
    <br>  
	<br>
	<br>
	<table width="611" border="0">  
      <tr>  
        <td height="29">
			<div class="text" style=" text-align:center;">
				授权给：<font color="#993300">   
			  <% =Application("aqjh_user")%>
			  </font>开放时间：<font color="#993300">   
			  <%=Application("aqjh_sn")%></font> 
			  访问统计：<font color="#993300">  
			  <%=Application("tuziji_visitor")%>			  
			  </font>  
			</div>
        </td>  
      </tr>
    </table> 
	<table width="611" border="0">  
      <tr>  
        <td height="29"><font color="#000000"><div class="text" style=" text-align:center;">CopyRight 2015-2015  备案号：粤ICP备15072406号</div></font>
        </td>  
      </tr>
    </table>  
  </form>  
</div>  
</body>  
