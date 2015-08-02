<%
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="DBQ="+server.mappath("../system/mxcz.asp")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
Conn.Open connstr
%>

<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("mian.mdb")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>
