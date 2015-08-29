<%
Set Conn1=Server.CreateObject("ADODB.Connection")
Connstr="DBQ="+server.mappath("Eline_Mima_GFSQ.asa")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
Conn1.Open connstr
%>
