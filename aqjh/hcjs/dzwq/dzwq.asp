<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
wqname=LCase(trim(Request.QueryString("wq")))
if wqname="" then
	Response.Write "<script language=javascript>{alert('提示：请正确选择你要修练的武器！');parent.history.go(-1);}</script>" 
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM x WHERE a='"&wqname&"'",conn,3,3
xlwp=rs("b")
money=rs("g")
tj=rs("h")
if isnull(xlwp) or xlwp="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：修练数据有问题不能操作\n找程序开发商联系!');</script>"
	response.end
end if
xadata=split(xlwp,"|")
xadata1=UBound(xadata)
rs.close
rs.open "select w6,w3,银两 from 用户 where 姓名='"&aqjh_name&"' and "&tj,conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不具备修练条件\n["&replace(tj,chr(39),"")&"]！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("银两")<money then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的银两不够["&money&"]两!');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你什么物品也没有!');</script>"
	response.end
end if
for ii=0 to xadata1-1
	xadata2=split(xadata(ii),":")
	mysl=clng(xadata2(1))
	myxlwp=trim(xadata2(0))
	duyao=abate(duyao,myxlwp,mysl)
next
zstemp=add(rs("w3"),wqname,1)
conn.execute "update 用户 set w6='"&duyao&"',w3='"&zstemp&"' where  姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("您[<%=wqname%>]修练完成！")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>