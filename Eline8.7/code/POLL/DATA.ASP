<%
Set conn=Server.CreateObject("ADODB.Connection")
Connstr="DBQ="+server.mappath("jx.mdb")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
conn.Open connstr
%>
