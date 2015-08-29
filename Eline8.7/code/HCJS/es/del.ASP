<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
id=trim(request.querystring("id"))
if not isnumeric(id) then Response.Redirect "../../error.asp?id=54"
myname=sjjh_name
if sjjh_grade<9 then
	Response.Write "<script Language=Javascript>alert('提示：你来这里作什么，想死呀！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from s where id=" & id & " and j>=0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：没有你要删除的物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
conn.execute "delete * from s where  id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：删除物品ID:"& id &"成功');location.href = 'javascript:history.go(-1)'</script>"
%>