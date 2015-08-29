<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<html>
<head>
<title>E线江湖鲜花店♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="huaimg/bg.gif" topmargin="0">
<table width="60%" border="0" height="205">
  <tr> 
    <td height="22" width="0%">&nbsp;</td>
    <td height="22" colspan="2"> 
      <div align="center"><img src="huaimg/top.gif" width="407" height="78"></div>
    </td>
    <td height="22" width="19%"><img src="huaimg/topright.gif" width="145" height="76"></td>
  </tr>
  <tr> 
    <td width="0%" rowspan="2">&nbsp;</td>
    <td width="53%" height="134" align="center" valign="top"> 
      <table width="100%" border="0" height="96">
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b where b='鲜花' order by b,h",conn,3,3
Response.Write "<tr>"
do while not rs.bof and not rs.eof
tj=replace(rs("l"),chr(39),"")
Response.Write "<TD  width='16%'><A href='#' onclick="&chr(34)&"showwp('"&rs("a")&"',"&rs("h")&",'"&rs("n")&"');"&chr(34)&" title='"&rs("c")&"'><IMG border=0 src='images/"&rs("i")&"'></A></TD>"
x=x+1
if x/3=int(x/3) then Response.Write "</tr><tr>"
rs.movenext
loop
Response.Write "</tr>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>      </table>
    </td>
    <td width="28%" height="134">
      <div align="right"><img src="huaimg/3.gif" width="174" height="237" border="0" usemap="#Map"></div>
    </td>
    <td width="19%" valign="bottom" rowspan="2"><img src="huaimg/4.gif" width="248" height="172"></td>
  </tr>
  <tr> 
    <td colspan="2" background="huaimg/l-22.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td width="0%" height="64">&nbsp;</td>
    <td colspan="3" height="64" id="zl">嗨，我是玛丽，欢迎来到鲜花小屋，点击即可获得信息。</td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="79,7,169,139" href="#" onclick=count() onmouseover=count()>
</map>
<script language=javascript>
function showwp(name,money,jyz){
	zl.innerHTML='<FORM action=buyhua.asp method=post name=form1><input type=hidden name=huaname value='+name+'>您选择的是：[<font color=blue>'+name+'</font>]，经营者为:'+jyz+',价值为：'+money+'两,要买<input type=text name=huasl size=5 maxLength=5 value="1">朵! <img src="huaimg/yes0.gif" onclick=yesno("'+name+'",document.form1.huasl.value) ></form> '
}
function yesno(name,sl) {
if(confirm("您现在将要购买的是 ["+name+"] 计："+sl+"朵 \n您现在可以点确定购买。")) {document.form1.submit();}
}

var dialogcount=10;
function count() {
	random=Math.random();
	if (random>0&&random<0.1&&dialogcount>0)
		zl.innerHTML="<font color=blue>我是玛丽，欢迎光临鲜花小屋，想买什么花？</font>";
	if (random>=0.1&&random<0.2&&dialogcount>0)
		zl.innerHTML="<font color=blue>看好了吗，买那一种花？</font>";
	if (random>=0.2&&random<0.3&&dialogcount>0)
		zl.innerHTML="<font color=blue>还不能选择吗，不会比考试难吧！</font>";
	if (random>=0.3&&random<0.4&&dialogcount>0)
		zl.innerHTML="<font color=blue>夏天选一只马蹄莲吧，听我的，没错的。</font>";
	if (random>=0.4&&random<0.5&&dialogcount>0)
		zl.innerHTML="<font color=blue>给自己爱的人选择一枝玫瑰是不错的选择。</font>";
	if (random>=0.5&&random<0.6&&dialogcount>0)
		zl.innerHTML="<font color=blue>炽热的爱情当然选择火烈鸟。</font>";
	if (random>=0.6&&random<0.7&&dialogcount>0)
		zl.innerHTML="<font color=blue>送给小MM当然是选择满天星。</font>";
	if (random>=0.7&&random<0.8&&dialogcount>0)
		zl.innerHTML="<font color=blue>大方得体当然是选择牡丹。</font>";
	if (random>=0.8&&random<0.9&&dialogcount>0)
		zl.innerHTML="<font color=blue>丝丝浪漫当然是选择蝴蝶兰。</font>";
	if (random>=0.9&&random<1&&dialogcount>0)
		zl.innerHTML="<font color=blue>爱情永远，当然是选择百合。</font>";
	if (dialogcount<1)
		zl.innerHTML="<font color=red>干吗老碰我，想泡我吗，没门。</font>";	
	dialogcount--;
}
</script>
</body>
</html>
