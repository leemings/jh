<%
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<head><link rel="stylesheet" href="css.css">
<title>♀wWw.happyjh.com♀</title>
<style>
body{
scrollbar-face-color:#bc8753;
scrollbar-shadow-color:#ffffff;
scrollbar-highlight-color:#ffffff;
scrollbar-3dlight-color:#bc8753;
scrollbar-darkshadow-color:#bc8753;
scrollbar-track-color:#dcbc9c;
scrollbar-arrow-color:#ffffff;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" text="#FFFFFF" leftmargin="5">
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td width="100%" height="28"> <div align="center"><font color="#CCCCFF" size="2"><strong>江湖牌桌</strong></font></div></td>
  </tr>
</table>
<table border="1" cellspacing="0" width="140" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr> 
    <td> <div align="center"><img src="dpk/1.GIF" width="21" height="21"><font color=yellow><font color=yellow><font color=yellow> 
        <input type=button value='扑克牌桌' onClick="window.open('dpk-xp.asp','f3')" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2">
        </font></font></font> <br></p>
        </div>
      <p align="center"> <font size="2"><img src="dpk/1.GIF"><font color=yellow><font color=yellow><font color=yellow> 
        <input type=button value='麻将牌桌' onClick="window.open('dmj-xp.asp','f3')" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button22">
        </font></font></font></font></p>
</td>
  </tr>
</table>
</body>
