<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link rel="stylesheet" type="text/css" href="cc.css">
</head>

<body background="pic/bg.jpg">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="96%" id="AutoNumber1" height="178">
  <tr>
    <td width="74%" height="63">
    <p align="center"><img border="0" src="pic/title2.jpg"></td>
  </tr>
  <tr>
    <td width="74%" height="117">　
    <table border=1 bgcolor="#FFFFFF" align=center width="93%" cellpadding="0" cellspacing="1" bordercolor="#FFCC00" height="51">
      
  <tr bgcolor=#74E76D bordercolor="#000000"> 
    <td height="24" bgcolor="#FFFFCC" bordercolor="#FF6600" width="15%"> 
      <div align="center"><font size="2" color="#FF0000">名妓</font></div>
    </td>
    <td height="24" bgcolor="#FFFFCC" bordercolor="#FF6600" width="9%"> 
      <div align="center"><font size="2" color="#FF0000">美貌</font></div>
    </td>
    <td height="24" width="19%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">聊天的价格</font></div>
    </td>
    <td height="24" width="14%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">增加的体力</font></div>
    </td>
     <td height="24" width="10%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">现有靓照</font></div>
    </td>
<td height="24" width="15%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">被嫖的次数</font></div>
    </td>
    <td height="24" width="64%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">操做</font></div>
    </td>
    
   

  </tr>
  <!--#include file="jiu.asp"-->
  <%
sql="SELECT * FROM 名妓"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor=#DEAD63> 
    <td bgcolor="#FFFFFF" width="15%" height="21"> 
      <center>
        <font size="2"><%=rs("姓名")%><span class="p9"><font color="#6666FF"></font></span> 
        </font> 
      </center>
      <font size="2"></font> 
    <td bgcolor="#FFFFFF" width="9%" height="21"> 
      <div align="center"><font size="2"><%=rs("美貌度")%></font> </div>
    
    <td bgcolor="#FFFFFF" width="19%" height="21"> 
      <div align="center"><font size="2"><%=rs("美貌度")*1.5%> </font></div>
    </td>
    <td bgcolor="#FFFFFF" width="14%" height="21"> 
      <div align="center"><font size="2"><%=rs("美貌度")/2%> </font></div>
    </td>
    <td bgcolor="#FFFFFF" width="10%" height="21"> 
      <div align="center"><font size="2"><%=rs("photo")%> </font></div>
    </td>
<td bgcolor="#FFFFFF" width="15%" height="21"> 
      <div align="center"><font size="2"><%=rs("time")%> </font></div>
    </td>

    <td bgcolor="#FFFFFF" width="40%" height="21"> 
      <center>
        <font size="2"><a href=admin.asp?id=<%=rs("id")%>><span class="calen-curr">站长操作</span></a></font> 
      </center>
    </td>
  </tr>
  <%
rs.movenext
loop
set rs=nothing
connt.close%>
</table>

    
    </td>
  </tr>
</table>


<SCRIPT language=JavaScript1.2>
var snowsrc="pic/1.gif"; var no = 18;
var ns4up = (document.layers) ? 1 : 0; var ie4up = (document.all) ? 1 : 0; var dx, xp, yp; var am, stx, sty;
var i, doc_width = 800, doc_height = 1200;
if (ns4up) {
doc_width = self.innerWidth;
doc_height = self.innerHeight;
} else if (ie4up) {
doc_width = document.body.clientWidth;
doc_height = document.body.clientHeight;
}
dx = new Array();
xp = new Array();
yp = new Array();
am = new Array();
stx = new Array();
sty = new Array();
for (i = 0; i < no; ++ i) {
dx[i] = 0;
xp[i] = Math.random()*(doc_width-50);
yp[i] = Math.random()*doc_height;
am[i] = Math.random()*20;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
if (ns4up) {
if (i == 0) {
document.write("<layer name=\"dot"+ i +"\" left=\"15\" top=\"15\" visibility=\"show\"><img src='"+snowsrc+"' border=\"0\"></a></layer>");
} else {
document.write("<layer name=\"dot"+ i +"\" left=\"15\" top=\"15\" visibility=\"show\"><img src='"+snowsrc+"' border=\"0\"></layer>");
}
} else if (ie4up) {
if (i == 0) {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><img src='"+snowsrc+"' border=\"0\"></a></div>");
} else {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><img src='"+snowsrc+"' border=\"0\"></div>");
}
}
}
function snowNS() {
for (i = 0; i < no; ++ i) {
yp[i] += sty[i];
if (yp[i] > doc_height-50) {
xp[i] = Math.random()*(doc_width-am[i]-30);
yp[i] = 0;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
doc_width = self.innerWidth;
doc_height = self.innerHeight;
}
dx[i] += stx[i];
document.layers["dot"+i].top = yp[i];
document.layers["dot"+i].left = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("snowNS()", 10);
}
function snowIE() {
for (i = 0; i < no; ++ i) {
yp[i] += sty[i];
if (yp[i] > doc_height-50) {
xp[i] = Math.random()*(doc_width-am[i]-30);
yp[i] = 0;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
doc_width = document.body.clientWidth;
doc_height = document.body.clientHeight;
}
dx[i] += stx[i];
document.all["dot"+i].style.pixelTop = yp[i];
document.all["dot"+i].style.pixelLeft = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("snowIE()", 10);
}
if (ns4up) {
snowNS();
} else if (ie4up) {
snowIE();
}
         </SCRIPT>


</body>

</html>