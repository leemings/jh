<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=trim(request("id"))
you=trim(request("you"))
if aqjh_grade=10 and instr(Application("aqjh_admin"),aqjh_name)<>0 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
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
