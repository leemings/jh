<!--#include file="FUN.INC"-->
<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="DBQ="+server.mappath("stock.mdb")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
Conn.Open connstr
conn1
%>