<%
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="DRIVER={Microsoft Access Driver (*.mdb)};DBQ="+server.mappath("../system/mxcz.asp")+";DefaultDir='';DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=1024;MaxScanRows=8;PageTimeout=15;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
Conn.Open connstr,"admin","ncadmin!ncadmin!"
%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>