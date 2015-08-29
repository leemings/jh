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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="SELECT * FROM b WHERE 物品名='" & name & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
if clng(Request.Form("wupinll"))<0 then Response.Redirect "../error.asp?id=471"
if clng(Request.Form("wupintl"))<0 then Response.Redirect "../error.asp?id=471"
name=Request.Form("wupinname")
a=clng(Request.Form("wupinll"))
b=clng(Request.Form("wupintl"))
wupinll=(a+b)*500
set rs=server.createobject("adodb.recordset")
sql="select * from b where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("a")=name
rs("b")=request.form("lx")
rs("f")=a
rs("g")=b
rs("h")=wupinll
rs.update
rs.close
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%>
<%else%>
<script language=vbscript>
MsgBox "仓库已经有这个物品了，请你添加点别的吧"
location.href = "javascript:history.back()"
</script>
<%end if%>
