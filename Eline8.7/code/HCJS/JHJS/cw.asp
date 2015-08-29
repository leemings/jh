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
<title>E线江湖宠物店♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="huaimg/bg.gif" topmargin="0">
<table width="84%" border="0" height="205">
  <tr> 
    <td height="31" width="1%">&nbsp;</td>
    <td height="31" colspan="2"> 
      <div align="center"><img src="huaimg/petbg.gif" width="188" height="35"></div>
    </td>
    <td height="31" width="46%"><img src="huaimg/pet00.gif" width="80" height="60"></td>
  </tr>
  <tr> 
    <td width="1%" rowspan="2">&nbsp;</td>
    <td width="19%" height="134" align="center" valign="top"> 
      <div align="left"><img src="huaimg/petseller.gif" width="122" height="180" usemap="#Map2" border="0"></div>
    </td>
    <td width="34%" height="134"> 
      <table width="100%" border="0" height="96">
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b where b='宠物' order by b,h",conn,3,3
Response.Write "<tr>"
do while not rs.bof and not rs.eof
tj=replace(rs("l"),chr(39),"")
Response.Write "<TD  width='16%'><A href='#' onclick="&chr(34)&"showwp('"&rs("a")&"',"&rs("h")&","&rs("m")&",'"&rs("c")&"');"&chr(34)&" title='"&rs("c")&"'><IMG border=0 src='images/"&rs("i")&"'></A></TD>"
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
    <td width="46%" valign="bottom" rowspan="2"><img src="huaimg/4.gif" width="248" height="172"></td>
  </tr>
  <tr> 
    <td colspan="2" background="huaimg/l-22.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1%" height="64">&nbsp;</td>
    <td colspan="3" height="64" id="zl">嗨，快把自己喜欢的小动物领回家吧。</td>
  </tr>
</table>
<script language=javascript>
function showwp(name,money,jb,sm){
	zl.innerHTML='<FORM action=buycw.asp method=post name=form1><input type=hidden name=huaname value='+name+'>[<font color=blue>'+name+'</font>]，('+sm+')价值为：'+money+'两,金币：'+jb+'个! <img src="huaimg/yes0.gif" onclick=yesno("'+name+'") ></form> '
}
function yesno(name) {
if(confirm("您现在将要购买的是 ["+name+"] \n您现在可以点确定购买。")) {document.form1.submit();}
}

var dialogcount=20;
function count() {
	random=Math.random();
	if (random>0&&random<0.1&&dialogcount>0)
		zl.innerHTML="<font color=blue>买了宠物可要好好养他们哟！</font>";
	if (random>=0.1&&random<0.2&&dialogcount>0)
		zl.innerHTML="<font color=blue>每一个宠物都有他们的特性。</font>";
	if (random>=0.2&&random<0.3&&dialogcount>0)
		zl.innerHTML="<font color=blue>宠物可以攻击的，攻击力在于你的培养！</font>";
	if (random>=0.3&&random<0.4&&dialogcount>0)
		zl.innerHTML="<font color=blue>如果你是会员，可以选择会员宠物。</font>";
	if (random>=0.4&&random<0.5&&dialogcount>0)
		zl.innerHTML="<font color=blue>对动物你要倍加爱护，否则他会死的。</font>";
	if (random>=0.5&&random<0.6&&dialogcount>0)
		zl.innerHTML="<font color=blue>宠物是每5小时照顾一次。</font>";
	if (random>=0.6&&random<0.7&&dialogcount>0)
		zl.innerHTML="<font color=blue>每次可以照顾4次小动物。</font>";
	if (random>=0.7&&random<0.8&&dialogcount>0)
		zl.innerHTML="<font color=blue>当小动物心情不好时他可不听话的。</font>";
	if (random>=0.8&&random<0.9&&dialogcount>0)
		zl.innerHTML="<font color=blue>买只小猫吧，看他多可爱。</font>";
	if (random>=0.9&&random<1&&dialogcount>0)
		zl.innerHTML="<font color=blue>还不能选择吗？</font>";
	if (dialogcount<1)
		zl.innerHTML="<font color=red><b>干吗老碰我，想泡我吗，没门。</b></font>";	
	dialogcount--;
}
</script>
<map name="Map2"> 
  <area shape="rect" coords="17,18,81,135" href="#" onclick=count() onmouseover=count()>
</map>
</body>
</html>
