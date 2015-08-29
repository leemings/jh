<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>

<html>

<head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt; color:'000000';
	scrollbar-face-color:'222222'; 
	scrollbar-shadow-color:'222222'; 
	scrollbar-highlight-color:'EFBA56'; 
	scrollbar-3dlight-color:'222222'; 
	scrollbar-darkshadow-color: 'AB7510'; 
	scrollbar-track-color:'FBEED7'; 
	scrollbar-arrow-color:'AB7510' ;

	}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'000000'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff4d4d;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: '291C03'; border: 1 solid '291C03'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; 
	BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-RIGHT-WIDTH: 1px; 
	BORDER-RIGHT-COLOR:'EBA82C'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: '333333'; 
	border: 333333px solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'F3CB81'; color: '83590C'; 
	border: 1px solid AB7510;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'EBA82C'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'EBA82C'}
</style>
<style>.skin0 {
	BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; VISIBILITY: hidden; BORDER-LEFT: 1px solid; WIDTH: 80px; CURSOR: default; LINE-HEIGHT: 20px; BORDER-BOTTOM: 1px solid; FONT-FAMILY: Verdana; POSITION: absolute; BACKGROUND-COLOR: white
}
.menuitems {
	PADDING-RIGHT: 10px; PADDING-LEFT: 15px
}
                </style>



<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>【<%=Application("sjjh_chatroomname")%>】</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}</script>

</head>

<body bgcolor="#000000" topmargin="0"><div align="center">
<center> 
<table width="654" height="172" border="0" align="center" cellpadding="0" cellspacing="0" background="bg.jpg">
  <tr>
      <td width="100%" colspan="3" background="k2.gif" height="18">　</td>
    </tr>
    <tr>
      
    <td width="18" rowspan="4" background="k3.gif" height="136">　</td>
      <td width="88%" align="center" height="26"><img border="0" src="45.gif"></td>
      
    <td width="18" background="k3.gif" height="136" rowspan="4">　</td>
    </tr>
    <tr>
      <td width="88%" height="17" valign="top">
        <img border="0" src="dan.jpg" width="618" height="225">
      </td>
    </tr>
    <tr>
      <td width="88%" height="145" valign="top">
        <table border="0" width="100%" bgcolor="#FFFFE1">

          <tr>
            <td width="100%" align="center"><hr size=1 color="009933"></td>
          </tr>

          <tr>
            <td width="100%" align="center">[ <font color="#FF3300">江 湖 当 铺</font> ]           

	当铺老板惊叫道： 你的可当物品有：   
</td>
          </tr>
          <tr>
            <td width="100%">        
<table border="0" cellspacing="0" cellpadding="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFE1" align="center" >
<tr>
<td width="75" >
<div align="center">物 品 名</div>
</td>
<td width="86" >
<div align="center">数量</div>
</td>
<td width="120" >
<div align="center">操作</div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w1,w2,w4,w5,w6,w7,w8,w9,w10 from 用户 where 姓名='"&sjjh_name&"'",conn
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>药 品 类</font></b></div></td>"
if not(isnull(rs("w1"))) then
call show(rs("w1"),"使用","w1")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>毒 药 类</font></b></div></td>"
if not(isnull(rs("w2"))) then
call show(rs("w2"),"下毒","w2")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>暗 器 类</font></b></div></td>"
if not(isnull(rs("w4"))) then
call show(rs("w4"),"投掷","w4")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>鲜 花 类</font></b></div></td>"
if not(isnull(rs("w7"))) then
call show(rs("w7"),"送花","w7")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>配 药 类</font></b></div></td>"
if not(isnull(rs("w8"))) then
	call show(rs("w8"),"配药","w8")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>魔 宝 类</font></b></div></td>"
if not(isnull(rs("w9"))) then
	call show(rs("w9"),"发射","w9")
end if

Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>秘 笈 类</font></b></div></td>"
if not(isnull(rs("w10"))) then
	call show(rs("w10"),"执行","w10")
end if

Response.Write "<tr></tr><tr></tr>"
rs.close
set rs=nothing
conn.close
set conn=nothing
sub show(data,lx,lb)
data1=split(data,";")
data2=UBound(data1)
for y=0 to data2-1
data3=split(data1(y),"|")
%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td width="75">
<div align="center">
<%=data3(0)%>
</div>
</td>
<td width="86">
<div align="center"><%=data3(1)%> </div>
</td>
<td width="120">
<div align="center"><a href="dan1.asp?lb=<%=lb%>&name=<%=data3(0)%>&sl=1">当1个</a>&nbsp;&nbsp;<a href="dan1.asp?lb=<%=lb%>&name=<%=data3(0)%>&sl=10">当10个</a></div>
</td>
</tr>
<%
erase data3
next
erase data1
end sub%>
</table>
            </td>
          </tr>
          <tr>
            <td width="100%">        
        

              　
            </td>
          </tr>
          <tr>
            <td width="100%" align="center">        
        

              选好要当的物品喔不要当错了
            </td>
          </tr>
        <td><hr size=1 color="009933">

          <tr>
            <td width="100%" align="center">『E线江湖』</td>
          </tr>
        </table>



  </center>

  </table>
  <div align="center">
    <center>
    <table width="654" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="100%" background="k2.gif" height="18">　</td>
      </tr>
    </table>
    </center>
  </div>

</body>







































