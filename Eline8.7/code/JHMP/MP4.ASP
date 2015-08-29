<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
id=trim(request("id"))
you=trim(request("you"))
if sjjh_grade=10 and instr(Application("sjjh_admin"),sjjh_name)<>0 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	conn.execute "update p set c=c-1 where a='" & id & "'"
	conn.execute "delete from 用户 where 姓名='" & you & "'"
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('门派：{"&id&"}的弟子["&you&"]删除成功！');location.href = 'javascript:history.back()';}</script>"
else
	Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');location.href = 'javascript:history.back()';}</script>"
end if
%>
