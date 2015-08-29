<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from b Where a='"&trim(Request.Form("a"))&"'",conn,2,2
if not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('提示：此物品已经存在，请换名！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
sqlstr="SELECT * FROM b "
rs.Open sqlstr,conn,1,2
rs.AddNew
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