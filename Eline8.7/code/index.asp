<bgsound src="wma/04.wma" loop="1">
<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Dim SplitReflashPage
Dim DoReflashPage
dim shuaxin_time
DoReflashPage=Application("sjjh_DoReflashPage")
shuaxin_time=10
ReflashTime=Now()
if (not isnull(session("ReflashTime"))) and cint(shuaxin_time)>0 and DoReflashPage then
	if DateDiff("s",session("ReflashTime"),Now())<cint(shuaxin_time) then
   	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��<b><font color=ff0000>"&shuaxin_time&"</font></b>��������ˢ�±�ҳ��<BR>���ڴ�ҳ�棬���Ժ򡭡�"
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
<title><%=Application("sjjh_chatroomname")%></title>
<LINK href="style.css" rel=stylesheet>
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
if(document.forms[0].username.value.length<=1){alert('���������Ѿ�����ɹ����û����ƣ�');document.forms[0].username.select();return false;}
else if(document.forms[0].username.value.indexOf("\'")!=-1||document.forms[0].username.value.indexOf("\"")!=-1){alert('����ʹ�÷Ƿ��ַ�������');document.forms[0].username.select();return false;}
else if(document.forms[0].password.value.length<6){alert('������̣�����֤�Ƿ�������ȷ');document.forms[0].password.select();return false;}
else if(document.forms[0].password.value.indexOf("\'")!=-1||document.forms[0].password.value.indexOf("\"")!=-1){alert('����ʹ�÷Ƿ��ַ�������');document.forms[0].password.select();return false;}
else {return true;}
}
</script>
</head>
<body onload="javascript:document.forms[0].username.focus();" bgcolor="<%=bgcolor%>" background="images/bg1.gif">
<div align=center> 
  <form action="check0.asp" method=post onsubmit='return(check());' name=form1>
    <table width="609" border=0>
      <tr> 
        <td align="center" valign="middle" width="131"><img src="images/STAMP.png" width="220" height="350"></td>
        <td align="center" valign="middle" width="450"> 
          <p><img src="images/poem3.gif" width="281" height="170"></p>
          <p><img src="images/partner.gif" width="450" height="60"></p>
        </td>
      </tr>
      <tr>
		<td valign="middle" align=center width="581" colspan="2">����<input type="text" name="name" value="" size=14 maxlength=14>
          ����  
          <input type="password" name="pass" value="" size=14 maxlength=14> 
          <input type="submit" value=" �� ¼ "> 
		  <script language=JavaScript src="online.asp"></script>
        </td> 
          
      </tr> 
       <tr ><td width="595" align="center" colspan="2">&nbsp;
		  <input onClick="window.open('yamen/read.htm','','width=400,height=502,menubar=no,top=100,left=200')" title=ע���ʺ� type=button value="ע  ��" name="button"> 
          <input onClick="window.open('yamen/disp.asp','casper','width=400,height=300,top=200,left=300')" title=�ʺŸ��� type=button value="��  ��" name="button">     
		  <input onClick="window.open('yamen/baoshi.asp','baoshi','width=400,height=320,top=200,left=300')" title=�� �� type=button value="��  ��" name="button">
          <input onClick="window.open('yamen/getid.asp','','width=400,height=250,menubar=no,top=200,left=300')" title=ȡ��ID type=button value="ȡ��ID" name="button">
          <input onClick="window.open('yamen/close.asp','close','width=400,height=320,top=200,left=300')" title=�����Ծ� type=button value="�����Ծ�" name="button">
          <input onClick="window.open('yamen/modify.asp','modify','width=400,height=350,top=200,left=300')" title=�������� type=button value="��������" name="button">     
          <input onClick="window.open('yamen/jiejiu.asp','jiejiu','width=400,height=350,top=200,left=300')" title=˯���Ծ� type=button value="˯���Ծ�" name="button">
        </td></tr>     
    </table>     
    �� 

    <table width="500" border="0">  
      <tr>  
        <td height="29"><font color="#CC0000">������á�</font>�����ֽ�����:����Ҵ��Ǽ��꣬���Ƿ񻹼ǵá���
        </td>  
      </tr>  
	  <tr>
	  <td height="29"><font color="#CC0000">���������������ո�վ�����ٸ�����Ա����������ϵվ��QQ865240608�����������н���Ŷ</font>
	  </tr>
    </table>  
    <br>  
	<br>
	<br>
	<table width="611" border="0">  
      <tr>  
        <td height="29">
			<div class="text" style=" text-align:center;">
				��Ȩ����<font color="#993300">   
			  <% =Application("sjjh_user")%>
			  </font>����ʱ�䣺<font color="#993300">   
			  <%=Application("sjjh_sn")%></font> 
			  ����ͳ�ƣ�<font color="#993300">  
			  <%=10240%>
			  </font>��Ա���ܣ�<font color="#993300">  
			  <%if Application("Ba_jxqy_fellow")=true then%>
			  ����
			  <%else%>
			  �ر�
			  <%end if%>
			  </font>  
			</div>
        </td>  
      </tr>
    </table> 
	<table width="611" border="0">  
      <tr>  
        <td height="29"><font color="#000000"><div class="text" style=" text-align:center;">CopyRight 2015-2015  �����ţ���ICP��15072406��</div></font>
        </td>  
      </tr>
    </table>  
  </form>  
</div>  
</body>  
