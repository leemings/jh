<%@ LANGUAGE=VBScript codepage ="936" %><%
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>特技列表</title>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" background="<%=chatimage%>" leftmargin="8" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" background="<%=chatimage%>" >
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#CCCCFF"><strong>特技列表</strong></font></div></td>
      </tr>
    </table>
<table width="100%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolorlight="#000000"
bordercolordark="#FFFFFF" bgcolor="#006699" background="<%=chatimage%>" >
<tr>
    <td><div align="center"><font color="#FFFFFF"><span style='font-size:9pt'>【</font></span><a href="stunt.asp" target="f3"><font color="#00FF99">连续特技</a></font><font color="#FFFFFF">】</font><br>    <font color="#FFFFFF"><span style='font-size:9pt'> 
        <br>
       <font color="#FFFFFF"><span style='font-size:9pt'>【</font></span></font><a href="compact.asp" target="f3"><font color="#00FF99">合体特技</a></font><font color="#FFFFFF">】</font></font><br>
        <br>
        <font color="#FFFFFF"><span style='font-size:9pt'>【</font></span></font><a href="dgjj.asp" target="f3"><font color="#00FF99">独孤九剑</font></a></font><font color="#FFFFFF">】</font></font> 
        <br>
        <br>
        <font color="#FFFFFF"><span style='font-size:9pt'>【</font></span></font><a href="zjjj.asp" target="f3"><font color="#00FF99">终级绝技</font></a></font><font color="#FFFFFF">】</font></font> 
        <br>
        <br>
        <font color="#FFFFFF"><span style='font-size:9pt'>【</font></span></font><font color="#000000"><a href="HYJZ.ASP" target="f3"><font color="#00FF99">会员绝招</font></a></font><font color="#FFFFFF">】</font></font> <br>
      </div></td>
</tr>
</table></body></html>