<%@ codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<head><link rel="stylesheet" href="css.css"></head>
<body bgcolor='4a79b5' background=../bg.gif>
<div align=center> 
<hr>
<p><br>
<br>
<img src="dpk/1.GIF" width="21" height="21"><font color=yellow><font color=yellow><font color=yellow> 
<input type=button value='ÆË¿ËÅÆ×À' onClick="window.open('dpk-xp.asp','f3')" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2">
</font></font></font> <br>
</p>
<p> <font size="2"><img src="../../JHIMG/SZ.GIF" width="20" height="19"><font color=yellow><font color=yellow><font color=yellow> 
<input type=button value='Âé½«ÅÆ×À' onClick="window.open('dmj-xp.asp','f3')" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button22">
</font></font></font></font></p>
<p>&nbsp; </p>
</div>
</body>