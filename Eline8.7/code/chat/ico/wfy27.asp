<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='广告'",conn,2,2
%>
<html>
<head>
<title>下图♀wWw.happyjh.com♀</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" text="#996600" link="#996600" vlink="#996600" alink="#996600" leftmargin="0" topmargin="0" bgproperties="fixed" marginwidth="0" marginheight="0" oncontextmenu=self.event.returnValue=false>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="">
        <tr> 
          <td width="34" height="28" background="wfy_i_c_g_b20.gif">&nbsp;</td>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="28" background="wfy_i_c_g_b21.gif"> 
                  <%=rs("c")%>
<%rs.close 
set rs=nothing
conn.close
set rs=nothing
%>

 </td>
              </tr>
            </table></td>
          <td width="34" height="28" background="wfy_i_c_g_b19.gif">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</html>
