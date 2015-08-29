<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
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
wupinname=trim(Request.Form("wupinname"))
wupinlx=trim(Request.Form("wupinlx"))
wupinnl=abs(trim(Request.Form("wupinnl")))
wupintl=abs(trim(Request.Form("wupintl")))
wupingj=abs(trim(Request.Form("wupingj")))
wupinsm=trim(Request.Form("wupinsm"))
wupinfy=abs(trim(Request.Form("wupinfy")))
wupinyl=abs(Request.Form("wupinyl"))
id=trim(Request.Form("wupinID"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.Execute "update b set a='"& wupinname &"',b='"& wupinlx &"',d="& wupinnl &",e="& wupintl &",f="& wupingj &",c='"& wupinsm &"',g="& wupinfy &",h="& wupinyl &" where id="&ID
Response.Redirect "../ok.asp?id=700"
%>