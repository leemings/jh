<%@ LANGUAGE=VBScript codepage ="936" %>

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
<title>抢劫银行</title>
</head>

<body bgcolor="#000000" leftmargin="0" topmargin="0">
<table width="90%" height="172" border="0" align="center" cellpadding="0" cellspacing="0" background="../../picc/bg.jpg">
  <tr>
      <td width="100%" colspan="3" background="k2.gif" height="18">　</td>
    </tr>
    <tr>
      
    <td width="2%" rowspan="4" background="k3.gif" height="136">　</td>
      <td width="88%" align="center" height="26">　</td>
      
    <td width="2%" background="k3.gif" height="136" rowspan="4">　</td>
    </tr>
    <tr>
      <td width="88%" height="17" valign="top">
        <table border="0" width="100%">
          <tr>
            <td width="20%"><img border="0" src="IN7.gif" width="95" height="130"></td>
            <td width="60%" valign="top" align="center">
              <img border="0" src="in9.jpg" width="327" height="156">
            </td>
            <td width="20%" align="right"><img border="0" src="IN8.gif" width="95" height="130"></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="88%" height="145" valign="top">
        <table border="0" width="100%" bgcolor="#FFFFE8">

          <tr>
            <td width="100%" align="center"><hr></td>
          </tr>

          <tr>
            <td width="100%" align="center">抢劫钱庄</td>
          </tr>
          <tr>
            <td width="100%">        

              <table border="0" width="100%" bordercolor="#666666">
                <tr>
                  <td width="100%">
                    <table border="1" width="100%">
                      <tr>
                        <td width="25%" align="center">钱庄</td>
                        <td width="25%" align="center">钱庄警戒</td>
                        <td width="25%" align="center">损失内力 体力</td>
                        <td width="25%" align="center">小心被抓</td>
                      </tr>
                       <!--#include file="loot3.asp"-->  
              <%  
sql="SELECT * FROM 抢劫"  
Set Rs=connt.Execute(sql) 
randomize
theb=int((rnd*10)+1) 
do while not rs.bof and not rs.eof  
%>  

                      <tr>
                        <td width="25%"><%=rs("默默姓名")%></td>
                        <td width="25%"><%=rs("默默爱美")*3%></td>
                        <td width="25%"><%=rs("默默爱美")*3%></td>
                        <td width="25%"><a href=loot2.asp?id=<%=rs("id")%>><span class="calen-curr">抢劫</span></a></td>
                      </tr>  <%  
rs.movenext  
loop  
set rs=nothing  
connt.close%>  

                    </table>
</td>
</tr>

                <tr>
                  <td width="100%">
                    
                  </td>
                </tr>
              </table>
            </td>
          </tr>
         <td><hr>

          <tr>
            <td width="100%" align="center">『E线江湖』</td>
          </tr>
        </table>



  </center>

  </table>
  <div align="center">
    <center>
    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="100%" background="k2.gif" height="18">　</td>
      </tr>
    </table>
    </center>
  </div>

</body>






