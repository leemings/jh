<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<LINK href="style1.css" rel=stylesheet>
<script language=javascript>
function check(){
if(document.forms[0].username.value.length<=1){alert('���������Ѿ�����ɹ����û����ƣ�');document.forms[0].username.select();return false;}
else if(document.forms[0].username.value.indexOf("\'")!=-1||document.forms[0].username.value.indexOf("\"")!=-1){alert('����ʹ�÷Ƿ��ַ�������');document.forms[0].username.select();return false;}
else if(document.forms[0].password.value.length<6){alert('������̣�����֤�Ƿ�������ȷ');document.forms[0].password.select();return false;}
else if(document.forms[0].password.value.indexOf("\'")!=-1||document.forms[0].password.value.indexOf("\"")!=-1){alert('����ʹ�÷Ƿ��ַ�������');document.forms[0].password.select();return false;}
else {return true;}
}
</script>
</head>
<body onload="javascript:document.forms[0].username.focus();" bgcolor="<%=bgcolor%>" background="<%=bgimage%>"><center>
  <form action="help1.asp" method=post onsubmit='return(check());' name=form1>
    <table width="398">
      <tr>
		<td valign="middle" align=center width="390">����<input type="text" name="username" value="" size=14 maxlength=14>
          ����  
          <input type="password" name="password" value="" size=14 maxlength=14> 
          <input type="submit" value="���ߴ���"> 
        </td> 
      </tr> 
    </table>  
  </form>  
</body>  
