<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"

%>
<html>
<body background="IMAGES/bg.jpg">
<div align="center"><a href="javascript:history.go(-1);"><font color="#0000FF">按这里返回</font></a></div>
<%
houseid=clng(trim(Request("houseid")))
if houseid<0 or houseid>5 then
	Response.Write "<script Language=Javascript>alert('提示：选择数据有错，请重新选择！');location.href = 'house.asp';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,1
if not(rs.Eof and rs.Bof) then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('提示：你已经购买了房屋了，要办理升级请去客厅操作！');location.href = 'house.asp';</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM housetype WHERE ht_序号=" & houseid,conn,1,1
xlwp=rs("ht_1级条件")
if isnull(xlwp) or xlwp="" or instr(xlwp,"|")=0 or instr(xlwp,";")=0 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：房屋数据有错不能操作\n请找程序开发商联系!');</script>"
	response.end
end if
xadata=split(xlwp,";")
xadata1=UBound(xadata)
rs.close
rs.open "SELECT w6,等级,会员等级,银两,金币 from [用户] WHERE 姓名='" & aqjh_name & "'",conn,1,3
duyao=rs("w6")
if isnull(duyao) or duyao="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你什么物品也没有!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
mysql=""
yin=0
for ii=0 to xadata1
	xadata2=split(xadata(ii),"|")
	mysl=clng(xadata2(1))
	myxlwp=trim(xadata2(0))
	select case myxlwp
	case "等级"
		if rs("等级")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：您的等级未到["& mysl &"]级，不能购买!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
	case "金币"
		if rs("金币")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：您的金币没有["& mysl &"]个，不能购买!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
		rs("金币")=rs("金币")-mysl
	case "银两"
		yin=yin+mysl
		if rs("银两")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：您的银两没有["& mysl &"]两，不能购买!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
		rs("银两")=rs("银两")-mysl
	case "会员等级"
		if rs("会员等级")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：您的会员等级未到["& mysl &"]级，不能购买!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
	case else
		if mywpsl(duyao,myxlwp)<mysl then
    		Set rs = Nothing
   			conn.Close
  			Set conn = Nothing
   			Response.Write "<script Language=Javascript>alert('提示：物品"& myxlwp &"数量不足？');</script>"
			response.end
		end if
		duyao=abate(duyao,myxlwp,mysl)
	end select
next
rs("w6")=duyao
rs.update
rs.close
rs.open "SELECT * FROM housetype WHERE ht_序号=" & houseid,conn,1,3
h_name=rs("ht_小区名")
h_nai=rs("ht_1级耐久度")
rs("ht_居民")=rs("ht_居民")+1
rs("ht_小区资产")=rs("ht_小区资产")+int(yin/10000)
rs.update
rs.close
rs.open "SELECT * FROM house",conn,1,3
rs.addnew
rs("h_小区名")=h_name
rs("h_拥有者")=aqjh_name
rs("h_拥有者2")="无"
rs("h_类型")=houseid
rs("h_级别")=1
rs("h_参观")=0
rs("h_购买时间")=now()
rs("h_访问时间")=now()
rs("h_回家次数")=0
rs("h_耐久度")=h_nai
rs("h_夫妻钱柜")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=""Javascript"">alert(""恭喜你，现在有了自的家了．以后可以取妻生子过快快乐乐的江湖生活！"");window.close();</script>"
%></html>
