<%@ LANGUAGE=VBScript codepage ="936" %>

<% 


dim conn,connpic,DBPath
dim rspic
dim user_id,sql
dim picid,cur,pics


Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
jiutian=Session("jiutian")
username=Session("sjjh_name")
dj=Session("sjjh_jhdj")
sex=Session("sjjh_xb")
if dj<2 then
	response.write "你还是江湖小辈，就想来这种地方.<a href=index.asp>返回妓院</a>"
	response.end
end if
if sex="女" then
	response.write "怡红院的姑娘不接待女的，请止步！<a href=index.asp>返回妓院</a>"
	response.end
end if
%>
<!--#include file="jiu.asp"-->
<%
sql="select * from 名妓 where 姓名='" &username& "'"
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
response.write "你有没有搞错呀，本院哪有这个姑娘呀，你到别家去赎身吧<a href=index.asp>返回妓院</a>"
	response.end
end if


%>

<!--#include file="connpic.asp"--><%



Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from pic where 姓名='"&username&"'"
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
    
    <form enctype="multipart/form-data" action="addpic.asp" method=post onsubmit="return mysubmit(this)">
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
  <tr>
    <td width="74%" height="58">上传几张漂亮得照片能帮你吸引更多得公子哥　
    <form>
      <table border="0" width="100%" height="13" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100%" height="10" colspan="2"><font color="#FF0000"><b>您已上传<%=rs.recordcount%>张相片</b></font></td>
        </tr>
        <%for i=1 to rs.recordcount%>
        <tr>
          <td width="55%" height="5">第<%=i%>张</td>
          <td width="45%" height="5">
            <p align="right"><a href="displaypic.asp?id=<%=rs("id")%>">[查看]</a><a href="delphoto.asp?id=<%=rs("id")%>&user_id=<%=rs("user_id")%>">[删除]</a></td>
        </tr>
        <%rs.movenext%>
        <%next%>
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