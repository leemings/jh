



<% Option Explicit
%><%
dim rs_lar,rs,jiutian,mykind,mypai,user,grade
dim sql
dim i,user_id
user_id=request("id")




%>






<!--#include file="connpic.asp"--><%


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from pic where user_id=" & user_id
rs.open sql,conn,3,2


%>


<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link rel="stylesheet" type="text/css" href="cc.css">

<script lanugage="JavaScript">
<!--
function pop(pageurl) {
  var
popwin=window.open(pageurl,"popWin","scrollbars=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,WIDTH=450,HEIGHT=180,left=150,top=150");
  return false;
}
//-->
</script>

<Script language="javascript">
function mysubmit(theform)
{
    if(theform.big.value=="")
    {
    alert("请点击浏览按钮，选择您要上传的jpg或gif文件!")
    theform.big.focus;
    return (false);
    }
    else
    {
    str= theform.big.value;
    strs=str.toLowerCase();
    lens=strs.length;
    extname=strs.substring(lens-4,lens);
    if(extname!=".jpg" && extname!=".gif")
    {
    alert("请选择jpg或gif文件!");
    return (false);
    }
    }
    return (true);
}
</script>

</head>

<body background="pic/bg.jpg">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="47%" id="AutoNumber1" height="232">
  <tr>
    <td width="26%" height="117" rowspan="2"><img border="0" src="pic/a1a.gif"></td>
    <td width="74%" height="59">　
    <form>
      <table border="0" width="100%" height="13" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100%" height="10" colspan="2"><font color="#FF0000"><b>您已上传<%=rs.recordcount%>张相片</b></font></td>
        </tr>
        <%for i=1 to rs.recordcount%>
        <tr>
          <td width="55%" height="5">第<%=i%>张</td>
          <td width="45%" height="5">
            <p align="right"><a href="displaypic.asp?id=<%=rs("id")%>">[查看]</a><a href="admindel.asp?id=<%=rs("id")%>&user_id=<%=rs("user_id")%>">[删除]</a></td>
        </tr>
        <%rs.movenext%>
        <%next%>
      </table>
      </form>

    
    </td>
  </tr>
  <tr>
    <td width="74%" height="58">　
    
    <!--#include file="conn.asp"-->
<%Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from 名妓  where id=" & user_id
rs.open sql,conn,3,2

if rs.eof or rs.bof then
response.write "你有没有搞错呀，本院哪有这个姑娘呀"
	response.end
set rs=nothing
conn.close
set rs=nothing
end if


%>

      <form enctype="multipart/form-data" action="adminadd.asp?id=<%=rs("id")%>" method=POST onsubmit="return mysubmit(this)">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td width="277"><font color="#FF0000"><b>相片上传</b></font></td>
        </tr>
        <tr>
          <td width="331">
            <p align="center">大相片&nbsp; <input type="file" name="big" size="20"></p> 
          </td> 
        </tr> 
        <tr> 
          <td width="337"> 
            <p align="center"><input type="submit" value="  上传  " name="B3">　</p>
          </td>
        </tr>
      </table>
      </form>

    
    
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