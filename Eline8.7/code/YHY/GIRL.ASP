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
if session("advtime")<>"" then
if session("advtime")>now()-0.006 then  Response.Redirect "../error.asp?id=490"
end if
session("advtime")=now()
jiutian=Session("jiutian")
nickname=Session("sjjh_name")
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

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="select * from 用户 where 姓名='" & nickname & "'"
set rs=conn.execute(sql)
yl=rs("银两") %>
<!--#include file="jiu.asp"-->
<% id=request("id")
sql="select * from 名妓 where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
response.write "你有没有搞错呀，本院哪有这个姑娘呀，是不是想姑娘想疯了"
	response.end
set rs=nothing
conn.close
set rs=nothing
end if
mingji=rs("姓名")
meimao=rs("美貌度")
if yl<meimao*1.5  then
	response.write "没钱也想来这种高级的地方呀，请止步！"
            conn.close
	response.end
set conn=nothing
set rs=nothing
end if

sql="update 用户 set 内力=内力+"&meimao&"/2,体力=体力-100 where 姓名='"&nickname&"' "
conn.execute sql
sql1="update 用户 set 银两=银两-"&meimao&"*1.5 where 姓名='"&nickname&"' "
conn.execute sql1

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from 名妓 where id=" & id,connt,3,2

rs("time")=rs("time")+1
rs.update

connt.close
conn.close
set rs=nothing
%>

<%




user_id=request("id")

Set connpic = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("Eline_yhy_pic.mdb")
connpic.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rspic = Server.CreateObject("ADODB.Recordset")
sql="select * from pic where user_id=" & user_id
rspic.open sql,connpic,1,1

if rspic.eof and rspic.bof then
   picid=1
   cur=1
else
   rspic.pagesize=1
   cur=request("cur")
   if cur="" or clng(cur)<1 then cur=1
   if clng(cur)>rspic.pagecount then cur=rspic.pagecount
   rspic.absolutepage=cur
   picid=rspic("id")
end if
   pics=rspic.recordcount
%>


<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link rel="stylesheet" type="text/css" href="cc.css">

<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>

</head>

<body background="pic/bg.jpg">
<!--------层----------->
<!--#include file="home.asp" -->
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#FFFFFF" valign="top"> 

<div id="layer1" style="position: absolute; left: 0; top:0; width: 500; height: 400; z-index: 2; visibility: visible">
  
</div>
<!--------------------->
<div id="layer2" style="position: absolute; left:0; top: 0; width:500; height: 400; z-index: 1; visibility: hidden">
  <table border="0" width="633" height="420" cellspacing="0" cellpadding="0">
    <tr bordercolor="#FFFFFF">
      <td width="496" valign="top" align="left" height="420"> 　
        <div align="center">
          <center>
            <table border="0" width="500" height="400" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
              <tr>
             <td width="100%" height="151" valign="top" align="center"> <a href="" onclick="return false" ><img name="big" src="display.asp?id=<%=picid%>" onclick="MM_showHideLayers('layer2','','hide');MM_showHideLayers('layer','','show');MM_showHideLayers('layer1','','show');return false;" border="0"></a>
              </td>
              </tr>
              </table>
          </center>
        </div>
      </td>
    </tr>
  </table>
</div>
<div id="layer" style="position: absolute; left: 0; top: 0; width:750; height: 400; z-index: 0; visibility: visible">
<!------------------------------------------------------------------------------------------------------------------------------>
<table border="0" width="750" cellspacing="0" cellpadding="0">
  <tr>

  </tr>
</table>
<!------------------------------------------------------------------------------------------------------------------------------>
<!------------------------------------------------------------------------------------------------------------------------------>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" id="AutoNumber1" height="400">
  <tr>
    <td width="26%" height="117"><img border="0" src="pic/a1a.gif"></td>
    <td width="74%" height="117">　
    <table border="0" width="91%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
              <tr>
                <td width="100%" bgcolor="#F1F1F1" valign="bottom" align="center">
                  <%if pics<>0 then%><a href="girl.asp?id=<%=user_id%>&cur=<%=cur-1%>" title="查看上一张相片">←</a><%end if%>
                  <%if cur=1 then%> <%if pics=0 then%> <font color="#FF9900">还没有上传相片</font>  
                  <%else%> <font color="#FF9900">有相片<font color="red"><%=pics%></font>张</font>  
                  <%end if%> <%else%> <font color="#FF9900">第<font color="red"><%=cur%></font>张</font>  
                  <%end if%> <%if pics<>0 then%><a href="girl.asp?id=<%=user_id%>&cur=<%=cur+1%>" title="查看下一张相片">→</a><%end if%>  
                </td>  
              </tr>  
              <tr>  
                  <td width="100%"> 
                  <p align="center"> 
                  <a onclick="return false" href="display.asp?id=<%=picid%>"><img name="small" border="0" src="display.asp?id=<%=picid%>" width="100" height="135" onclick="MM_showHideLayers('layer2','','show')"></a>  
                  </td>  
              </tr>  
              
              
            </table>  

    
    <p align="center"><img border="0" src="pic/na.jpg" align="center"></td>
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