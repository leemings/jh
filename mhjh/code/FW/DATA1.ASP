<%
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="DRIVER={Microsoft Access Driver (*.mdb)};DBQ="+server.mappath("../21ex/fangzi.asp")+";DefaultDir='';DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=15;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
Conn.Open connstr
%>
