<%
Response.Write "<script Language=Javascript>parent.fs.location.href = 'MINE.asp';</script>"
%><html>
<head>
<title>深山采矿</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
  <table width="100%" border="1">
    <tr> 
      <td width="21%" height="19"> 
        <div align="center"><font size="-1">矿石数：</font></div>
      </td>
      <td width="14%" height="19" id="fm1"> 
        <div align="center"><font size="-1"></font></div>
      </td>
      <td width="16%" height="19"> 
        <div align="center"><font size="-1">矿石数:</font></div>
      </td>
      <td width="13%" height="19" id="fm2"> 
        <div align="center"><font size="-1"></font></div>
      </td>
      <td width="15%" height="19"> 
        <div align="center"><font size="-1">操作数:</font></div>
      </td>
      <td width="7%" height="19" id="fm3"> 
        <div align="center"><font size="-1"></font></div>
      </td>
      <td width="14%" height="19" > 
<FORM action="MINEjl.asp" method=POST name="shiform" target="ig">
<INPUT name="fmsl" type=hidden value=0> 
<div align="center"><font size="-1"><a href="#" onclick="document.shiform.submit()">保存</a></font></div>
      </td>
    </tr>
  </table></FORM></body></html>