<%
Response.Write "<script Language=Javascript>parent.fs.location.href = 'ice.asp';</script>"
%><html>
<head>
<title>采冰</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font size="-1">E线江湖</font> </div>
<FORM action="icejl.asp" method=POST name="shiform" target="cz">
  <table width="100%" border="1" cellspacing="3" cellpadding="3">
    <tr> 
      <td width="13%" height="8"> 
        <div align="center"><font size="-1">冰水数：</font></div>
      </td>
    </tr>
    <tr> 
      <td width="13%" height="9" id="fm1"> 
        <div align="center"><font size="-1"></font></div>
      </td>
    </tr>
    <tr> 
      <td width="13%" height="19"> 
        <div align="center"><font size="-1">采冰数:</font></div>
      </td>
    </tr>
    <tr> 
      <td width="13%" height="19" id="fm2"> 
        <div align="center" ><font size="-1"></font></div>
      </td>
    </tr>
    <tr> 
      <td width="13%" height="19"> 
        <div align="center"><font size="-1">操作数：</font></div>
      </td>
    </tr>
    <tr> 
      <td width="13%" height="19" id="fm3"> 
        <div align="center" ><font size="-1"></font></div>
      </td>
    </tr>
    <tr> 
      <td width="13%" height="19"> 
        <div align="center"> <font size="-1"> 
          <input name="fmsl" type=hidden value=0>
          <a href="#" onClick="document.shiform.submit()">保存</a></font> </div>
      </td>
    </tr>
  </table>
</FORM>  
</body>
</html>