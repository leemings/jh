<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<LINK href="style.css" rel=stylesheet>
<script language=javascript>

function getOs() 
{ 
    var OsObject = ""; 
   if(isChrome=navigator.userAgent.indexOf("Chrome")!=-1){ 
        return "<audio src='/mid/bg.mp3' type='audio/mp3' autoplay='autoplay' hidden='true'></audio> "; 
   }
   else if(isIE = navigator.userAgent.indexOf("MSIE")!=-1) { 
//        return "<embed autostart='true' loop='-1' controls='ControlPanel' width='0' height='0' src='/mid/bg.mp3'/>"; 
		return "<bgsound src='mid/bg.mp3'>";
   }
   else {
	   return "<embed autostart='true' loop='-1' controls='ControlPanel' width='0' height='0' src='/mid/bg.mp3'/>"; 
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
		<td valign="middle" align=center width="581" colspan="2">����<input type="text" name="username" value="" size=14 maxlength=14>
          ����  
          <input type="password" name="password" value="" size=14 maxlength=14> 
          <input type="submit" value=" �� ¼ "> 
          <script language=javascript src="onlinelist.asp"></script> 
        </td> 
          
      </tr> 
       <tr ><td width="595" align="center" colspan="2">&nbsp;<input onClick="javascript:window.open('regstep1.asp','register','left=200,top=100,width=500,height=360,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=ע���ʺ� type=button value="ע  ��" name="button"> 
          <input onClick="javascript:window.open('relive.asp','relive',' width=300,height=210,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=�ʺŸ��� type=button value="��  ��" name="button">     
          <input onClick="javascript:window.open('suicide.asp','suicide',' width=300,height=350,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=�� ɱ type=button value="��  ɱ" name="button">     
          <input onClick="javascript:window.open('zj/help.asp','readme',' width=500,height=300,left=200,top=100,status=no,toolbars=yes,menubars=yes,scrollbars=yes,resize=no')" title=�������� type=button value="�� ��" name="button">     
          <input onClick="javascript:window.open('faq.htm','faq',' width=500,height=300,left=100,top=100,status=no,toolbars=yes,menubars=yes,scrollbars=yes,resize=no')" title=������ type=button value="������" name="button">     
                    <input onClick="javascript:window.open('chgpw.asp','chgpassword',' width=300,height=300,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=�������� type=button value="��������" name="button">     
          <input onClick="javascript:window.open('chgdatu1.asp','chgdatum',' width=500,height=300,left=200,top=100,status=no,toolbars=no,menubars=no,scrollbars=no,resize=no')" title=���ϸ��� type=button value="���ϸ���" name="button">     
        </td></tr>     
    </table>     
    �� 

    <table width="611" border="0">  
      <tr>  
        <td height="29"><font color="#CC0000">������á�</font>�����ֽ�����:����Ҵ��Ǽ��꣬���Ƿ񻹼ǵá���
		
          <!--
		  <table b<!-- order="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>����λ��</b></font></td>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>����λ��</b></font></td>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>����λ��</b></font></td>
              <td width="17%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>����λ��</b></font></td>
              <td width="16%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>����λ��</b></font></td>
              <td width="16%" bgcolor="#0000FF" align="center"><font color="#FFFFFF"><b>����λ��</b></font></td>
            </tr>
          </table> 
		  -->
        </td>  
      </tr>  
	  <tr>
	  <td height="29"><font color="#CC0000">��ף��վͨ��ͨ�Ź������ˣ��ֶ��⿪�ţ���ӭǰ����ˣ�������������ϵվ��qq865240608</font>
	  </tr>
    </table>  
    <br>  
	<br>
	<br>
    <table width="100%" border="0" align="center">       
      <tr> 
        <td align="center" valign="middle">
          ��Ȩ����<font color="#993300">   
          <% response.write Application("Ba_jxqy_userright")%>
          </font>����ʱ�䣺<font color="#993300">   
          <%=Application("Ba_jxqy_opendata")%></font> 
          ����ͳ�ƣ�<font color="#993300">  
          <%=Application("Ba_jxqy_visitor")%>
          </font>��Ա���ܣ�<font color="#993300">  
          <%if Application("Ba_jxqy_fellow")=true then%>
          ����
          <%else%>
          �ر�
          <%end if%>
          </font>  
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
