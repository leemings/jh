<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
id=abs(trim(Request.Form("wupinid")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM b where id="&ID
rs.Open sqlstr,conn,1,2
rs("a")=trim(Request.Form("a"))
rs("b")=trim(Request.Form("b"))
rs("c")=trim(Request.Form("c"))
rs("d")=int(abs((Request.Form("d"))))
rs("e")=int(abs(Request.Form("e")))
rs("f")=int(abs(Request.Form("f")))
rs("g")=int(abs(Request.Form("g")))
rs("h")=int(abs(Request.Form("h")))
rs("i")=trim(Request.Form("i"))
rs("j")=int(abs(Request.Form("j")))
rs("k")=trim(Request.Form("k"))
rs("l")=trim(Request.Form("l"))
rs("m")=int(abs(Request.Form("m")))
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
'conn.Execute "update b set a='"& wupinname &"',b='"& wupinlx &"',d="& wupinnl &",e="& wupintl &",f="& wupingj &",c='"& wupinsm &"',g="& wupinfy &",h="& wupinyl &" where id="&ID
Response.Redirect "../ok.asp?id=700"
%>