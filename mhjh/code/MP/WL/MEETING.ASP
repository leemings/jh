<%
username=session("yx8_mhjh_username")
if username=""  then Response.Redirect "../error.asp?id=016"%>
<html>

<head>
<title>武林大会</title>

<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false>

<table width="551" border="0" cellspacing="0" cellpadding="0" align="center"
height="5">
<tr>
<td width="551"  height="1">
<table border="0" height="24" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="91" height="26"><font size="2">&nbsp; <font
color="#FFFFFF"></font><font size="2"><font color="#ffffff"
size="2"></font></font></font></td>
<td width="475" height="26">

<div align="center"> <img src="images/download.gif" width="142" height="24">
</div>

<div align="center"> 　
</div>
</td>
<td width="104">
<div align="center">
</div>
</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
</td>
</tr>
</table>
<table width="732" border="0" cellspacing="0" cellpadding="0" align="center" height="238">
<tr>
<td width="17" background="../jh/tiao10.gif" height="245">　</td>
<td width="47" valign="top" height="245">

</td>
<td valign="top" width="776" height="245">
<div align="center">
<table width="744" border="0" cellspacing="0" cellpadding="0" height="234">
<tr valign="top" align="center">
<td height="180" width="744">
<table width="738" border="0" cellspacing="0" cellpadding="0" height="141">
<tr>
<td width="11"><img src="images/rightct3.gif" width="11" height="13"></td>
<td background="images/rightct4.gif" height="13" class="td01" width="712"><img src="../yamen/image/empty.gif" width="1" height="1"></td>
<td width="9"><img src="images/rightct1.gif" width="11" height="13"></td>
</tr>
<tr>
<td class="td" background="images/rightct8.gif" width="11"><img src="../yamen/image/empty.gif" width="1" height="1">
</td>
<td valign="top" width="712">
<table width="741" border="0" cellspacing="0" cellpadding="0" height="125">
<tr>
<td align="center" height="35" width="96"><font color="#FF0000" size="4">国家</font></td>
<td align="center" height="35" colspan="2" width="253"><img src="images/gold.gif" width="120" height="40"></td>
<td height="35" align="center" colspan="2" width="270"><img src="images/silver.gif" width="120" height="40"></td>
<td height="35" align="center" colspan="2" width="218"><img src="images/copper.gif" width="120" height="40"></td>
</tr>
<tr>
  <%   
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
Set conn1=Server.CreateObject("ADODB.CONNECTION")   
  conn1.open connstr1   
set rs=server.createobject("adodb.recordset")   
						sql="select * from gold where 国家='齐'"   
						rs.open sql,conn1,1   
						%>  
<td width="26" height="22" align="right"><font color="#FFFF00">齐国</font></td>           
<td width="214" height="22" align="left">君主： <%=rs("姓名")%>           
</td>           
<td width="89" height="22" align="left"><a href='fight.asp?typename=gold&id=齐'><img src="images/fight.gif" width="80" height="30" border="0"></a>           
<%rs.movenext%>           
</td>           
<td width="181" height="22" align="left">丞相：            
                          <%      
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"      
Set conn1=Server.CreateObject("ADODB.CONNECTION")      
  conn1.open connstr1      
set rs1=server.createobject("adodb.recordset")      
						sql1="select * from silver where 国家='齐'"      
						rs1.open sql1,conn1,1      
						%>            
  <%=rs1("姓名")%>           
</td>           
<td width="101" height="22" align="left"><a href='fight.asp?typename=silver&id=齐'>
<img src="images/fight.gif" width="80" height="30" border="0"></a>           
<%rs1.movenext%>           
</td>           
<%connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
Set conn1=Server.CreateObject("ADODB.CONNECTION")     
  conn1.open connstr1     
set rs2=server.createobject("adodb.recordset")     
						sql2="select * from copper where 国家='齐'"     
						rs2.open sql2,conn1,1%>          
<td width="163" height="22" align="left">元帅：<%=rs2("姓名")%>               
</td>           
<td width="79" height="22" align="left"><a href='fight.asp?typename=copper&id=齐'>
<img src="images/fight.gif" width="80" height="30" border="0"></a>           
<%rs2.movenext%>           
</td>           
</tr>           
<tr>           
<td width="26" height="23" align="right"><font color="#FFFF00">楚国
</font>           
</td>           
  <%   
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
Set conn1=Server.CreateObject("ADODB.CONNECTION")   
  conn1.open connstr1   
set rs=server.createobject("adodb.recordset")   
						sql="select * from gold where 国家='楚'"   
						rs.open sql,conn1,1   
						%>  

<td width="214" height="23" align="left">君主：<%=rs("姓名")%>           
</td>           
<td height="23" align="left" width="89"><a href='fight.asp?typename=gold&id=楚'><img src="images/fight.gif" width="80" height="30" border="0"></a>           
<%rs.movenext%>           
</td>           
                         <%      
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"      
Set conn1=Server.CreateObject("ADODB.CONNECTION")      
  conn1.open connstr1      
set rs1=server.createobject("adodb.recordset")      
						sql1="select * from silver where 国家='楚'"      
						rs1.open sql1,conn1,1      
						%>            

<td width="181" height="23" align="left">丞相：<%=rs1("姓名")%>           
</td>           
<td width="101" height="23" align="left"><a href='fight.asp?typename=silver&id=楚'>
<img src="images/fight.gif" width="80" height="30" border="0"></a>           
<%rs1.movenext%>           
</td>           
<%connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
Set conn1=Server.CreateObject("ADODB.CONNECTION")     
  conn1.open connstr1     
set rs2=server.createobject("adodb.recordset")     
						sql2="select * from copper where 国家='楚'"     
						rs2.open sql2,conn1,1%>          

<td width="163" height="23" align="left">元帅：<%=rs2("姓名")%>           
</td>           
<td width="79" height="23" align="left"><a href='fight.asp?typename=copper&id=楚'>
<img src="images/fight.gif" width="80" height="30" border="0"></a>           
<%rs2.movenext%>           
</td>           
</tr>           
<tr>           
<td width="26" height="22" align="right"><font color="#FFFF00">燕国</font></td>           
 <% connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
Set conn1=Server.CreateObject("ADODB.CONNECTION")   
  conn1.open connstr1   
set rs=server.createobject("adodb.recordset")   
						sql="select * from gold where 国家='燕'"   
						rs.open sql,conn1,1%>  
<td width="214" height="22" align="left">君主：<%=rs("姓名")%></td>           
<td height="22" align="left" width="89"><a href='fight.asp?typename=gold&id=燕'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
                         <%      
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"      
Set conn1=Server.CreateObject("ADODB.CONNECTION")      
  conn1.open connstr1      
set rs1=server.createobject("adodb.recordset")      
						sql1="select * from silver where 国家='燕'"      
						rs1.open sql1,conn1,1      
						%>            
<td width="181" height="22" align="left">丞相：<%=rs1("姓名")%></td>           
<td width="101" height="22" align="left"><a href='fight.asp?typename=silver&id=燕'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
<%connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
Set conn1=Server.CreateObject("ADODB.CONNECTION")     
  conn1.open connstr1     
set rs2=server.createobject("adodb.recordset")     
						sql2="select * from copper where 国家='燕'"     
						rs2.open sql2,conn1,1%>          

<td width="163" height="22" align="left">元帅：<%=rs2("姓名")%></td>           
<td width="79" height="22" align="left"><a href='fight.asp?typename=copper&id=燕'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
</tr>           
<tr>           
<td width="26" height="22" align="right"><font color="#FFFF00">韩国</font></td>           
 <% connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
Set conn1=Server.CreateObject("ADODB.CONNECTION")   
  conn1.open connstr1   
set rs=server.createobject("adodb.recordset")   
						sql="select * from gold where 国家='韩'"   
						rs.open sql,conn1,1%>  

<td width="214" height="22" align="left">君主：<%=rs("姓名")%></td>           
<td height="22" align="left" width="89"><a href='fight.asp?typename=gold&id=韩'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
                         <%      
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"      
Set conn1=Server.CreateObject("ADODB.CONNECTION")      
  conn1.open connstr1      
set rs1=server.createobject("adodb.recordset")      
						sql1="select * from silver where 国家='韩'"      
						rs1.open sql1,conn1,1      
						%>            

<td width="181" height="22" align="left">丞相：<%=rs1("姓名")%></td>           
<td width="101" height="22" align="left"><a href='fight.asp?typename=silver&id=韩'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
<%connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
Set conn1=Server.CreateObject("ADODB.CONNECTION")     
  conn1.open connstr1     
set rs2=server.createobject("adodb.recordset")     
						sql2="select * from copper where 国家='韩'"     
						rs2.open sql2,conn1,1%>          

<td width="163" height="22" align="left">元帅：<%=rs2("姓名")%></td>           
<td width="79" height="22" align="left"><a href='fight.asp?typename=copper&id=韩'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
</tr>           
<tr>           
<td width="26" height="22" align="right"><font color="#FFFF00">赵国</font></td>           
 <% connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
Set conn1=Server.CreateObject("ADODB.CONNECTION")   
  conn1.open connstr1   
set rs=server.createobject("adodb.recordset")   
						sql="select * from gold where 国家='赵'"   
						rs.open sql,conn1,1%>  

<td width="214" height="22" align="left">君主：<%=rs("姓名")%></td>           
<td height="22" align="left" width="89"><a href='fight.asp?typename=gold&id=赵'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
                         <%      
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"      
Set conn1=Server.CreateObject("ADODB.CONNECTION")      
  conn1.open connstr1      
set rs1=server.createobject("adodb.recordset")      
						sql1="select * from silver where 国家='赵'"      
						rs1.open sql1,conn1,1      
						%>            

<td width="181" height="22" align="left">丞相：<%=rs1("姓名")%></td>           
<td width="101" height="22" align="left"><a href='fight.asp?typename=silver&id=赵'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
<%connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
Set conn1=Server.CreateObject("ADODB.CONNECTION")     
  conn1.open connstr1     
set rs2=server.createobject("adodb.recordset")     
						sql2="select * from copper where 国家='赵'"     
						rs2.open sql2,conn1,1%>          

<td width="163" height="22" align="left">元帅：<%=rs2("姓名")%></td>           
<td width="79" height="22" align="left"><a href='fight.asp?typename=copper&id=赵'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
</tr>           
<tr>           
<td width="26" height="22" align="right"><font color="#FFFF00">魏国</font></td>           
 <% connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
Set conn1=Server.CreateObject("ADODB.CONNECTION")   
  conn1.open connstr1   
set rs=server.createobject("adodb.recordset")   
						sql="select * from gold where 国家='魏'"   
						rs.open sql,conn1,1%>  

<td width="214" height="22" align="left">君主：<%=rs("姓名")%></td>           
<td height="22" align="left" width="89"><a href='fight.asp?typename=gold&id=魏'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
                         <%      
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"      
Set conn1=Server.CreateObject("ADODB.CONNECTION")      
  conn1.open connstr1      
set rs1=server.createobject("adodb.recordset")      
						sql1="select * from silver where 国家='魏'"      
						rs1.open sql1,conn1,1      
						%>            

<td width="181" height="22" align="left">丞相：<%=rs1("姓名")%></td>           
<td width="101" height="22" align="left"><a href='fight.asp?typename=silver&id=魏'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
<%connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
Set conn1=Server.CreateObject("ADODB.CONNECTION")     
  conn1.open connstr1     
set rs2=server.createobject("adodb.recordset")     
						sql2="select * from copper where 国家='魏'"     
						rs2.open sql2,conn1,1%>          

<td width="163" height="22" align="left">元帅：<%=rs2("姓名")%></td>           
<td width="79" height="22" align="left"><a href='fight.asp?typename=copper&id=魏'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
</tr>           
<tr>           
<td width="26" height="22" align="right"><font color="#FFFF00">秦国</font></td>           
 <% connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
Set conn1=Server.CreateObject("ADODB.CONNECTION")   
  conn1.open connstr1   
set rs=server.createobject("adodb.recordset")   
						sql="select * from gold where 国家='秦'"   
						rs.open sql,conn1,1%>  

<td width="214" height="22" align="left">君主：<%=rs("姓名")%></td>           
<td height="22" align="left" width="89"><a href='fight.asp?typename=gold&id=秦'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
                         <%      
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"      
Set conn1=Server.CreateObject("ADODB.CONNECTION")      
  conn1.open connstr1      
set rs1=server.createobject("adodb.recordset")      
						sql1="select * from silver where 国家='秦'"      
						rs1.open sql1,conn1,1      
						%>            

<td width="181" height="22" align="left">丞相：<%=rs1("姓名")%></td>           
<td width="101" height="22" align="left"><a href='fight.asp?typename=silver&id=秦'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
<%connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
Set conn1=Server.CreateObject("ADODB.CONNECTION")     
  conn1.open connstr1     
set rs2=server.createobject("adodb.recordset")     
						sql2="select * from copper where 国家='秦'"     
						rs2.open sql2,conn1,1%>          

<td width="163" height="22" align="left">元帅：<%=rs2("姓名")%></td>           
<td width="79" height="22" align="left"><a href='fight.asp?typename=copper&id=秦'>
<img src="images/fight.gif" width="80" height="30" border="0"></a></td>           
</tr>           
</table>           
</td>           
<td class="td" background="images/rightct08.gif" width="9">　</td>           
</tr>           
<tr>           
<td height="13" width="11"><img src="images/rightct9.gif" width="11" height="10"></td>           
<td height="13" class="td01" background="images/rightct10.gif" width="712"><img src="../yamen/image/empty.gif" width="1" height="1"></td>           
<td width="9"><img src="images/rightct11.gif" width="11" height="10"></td>           
</tr>           
</table>           
<br>           
<table width="80%" border="0" cellspacing="0" cellpadding="0" height="49">           
<tr>           
<td height="60"><font color="#FF0000">江湖提醒</font>： 
夺取职位成功，自动获得这个国家的职位。如果失败，将要损失不少体力和内力,(<font color="#FFFF00"><span lang="zh-cn">切记：参与了就有损失，不管你是否有资格
；争夺成功后获得的积分等状态完全是从上任职位获得者那里争夺来的，并非系统赋予，也就是说，你加积分，就有人要减积分！</span></font>)。决定胜负的主要参数的攻击、防御、内力，但这三种状态并非简单的比较多少，而是一种综合参数。具体怎么比较，自己琢磨。^_^</td>           
</tr>           
</table>           
</td>           
</tr>           
</table>           
</div>           
</td>           
<td width="67" background="../jh/tiao10.gif" height="245">　</td>           
</tr>           
</table>           
           
<div align="center">           
</div>           
           
</body>           
           
</html>