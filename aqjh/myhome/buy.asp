<%id=clng(trim(request("id")))
name=trim(request("name"))
tj=trim(request("tj"))
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>购买小区</title>
<script language="JavaScript" type="text/JavaScript">
document.onmousedown=click;
function click(){if(event.button==2){alert("欢迎来到爱情江湖！");}}
</script>
<LINK href="style.css" rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0">
<table width="75%" border="0" cellspacing="0">
  <tr>
    <td><img src="pic/bk1.jpg" width="635" height="397" border="0" usemap="#Map">
      <div id="Layer2" style="position:absolute; width:464px; height:106px; z-index:2; left: 112px; top: 275px;">购买条件:<br>
        <font color="#0000FF"><%=replace(replace(tj,";","&nbsp;&nbsp;"),"|",":")%><br>
        <br>
        <br>
        </font> 
        <table width="83%" border="0" align="center">
          <tr>
            <td  align="right"> 
              <form name="form" method="post" action="buyok.asp">
                <input name="houseid" type="hidden" id="houseid" value="<%=id%>">
                <input type="submit" name="Submit" value="确 定 购 买">
              </form></td>
          </tr>
        </table>
        <font color="#0000FF"> </font></div>
      <div id="Layer1" style="position:absolute; width:182px; height:164px; z-index:1; left: 421px; top: 49px;"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="15" align="center"><%=name%> </td>
          </tr>
          <tr>
            <td height="128" align="left" valign="top"><img src="pic/hut<%=id%>1.jpg" width="185" height="125"></td>
          </tr>
          <tr>
            <td height="15" align="center"><font color="#0000FF"><%=name%> 1级</font></td>
          </tr>
        </table>
      </div></td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="25,25,91,52" href="javascript:history.go(-1);">
</map>
</body>
</html>
