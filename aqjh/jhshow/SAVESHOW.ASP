
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0


aqjh_name=Session("aqjh_name")
if aqjh_name="" then 
	Response.Redirect "login.asp"
	Response.End
end if
data=trim(Request("data"))
showdata=Split(data, "|")
if UBound(showdata)<>24 then
	Response.Write "<script language=JavaScript>{alert('提示：提交数据有问题，无法保存,请初始化数据重来！');window.close();}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
for i=0 to UBound (showdata)
xx=0
select case i
case 6
	if showdata(i)="14" or showdata(i)="7" then
		xx=1
	end if
case 7
	if showdata(i)="13" or showdata(i)="6" then
		xx=1
	end if
case 8
	if showdata(i)="12" or showdata(i)="5" then
		xx=1
	end if
case 10
	if showdata(i)="11" or showdata(i)="4" then
		xx=1
	end if
case 12
	if showdata(i)="10" or showdata(i)="3" then
		xx=1
	end if
case 13
	if showdata(i)="9" or showdata(i)="2" then
		xx=1
	end if
case 17
	if showdata(i)="8" or showdata(i)="1" then
		xx=1
	end if
case else
	xx=0
end select
if showdata(i)<>"" and xx<>1 then
	idsql="select * from data where i="&clng(showdata(i))
	rs.Open idsql, conn, 1,1
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：资料在系统库中不存在，请清空资料重新装备!');window.close();}</script>"
		Response.End
	end if
	id=rs("id")
	name=rs("b")
	spjb=rs("f")
	rs.close
	idsql="select * from usesp where sa='"& aqjh_name &"' and sb="& id
	rs.Open idsql, conn, 1,1
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{if(confirm('提示:["& name &" 价值:"& spjb &"金币]您还没有购买，现在购买吗？')){location.href='buy.asp?id="& id &"&sl=1';}}</script>"
		Response.End
	end if
	rs.close
end if
next 
idsql="select * from use where a='" & aqjh_name & "'"
rs.Open idsql, conn, 1,1
if rs("i")>=3 and rs("d")=false then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：您非会员，仅可以保存3张像册，会员可以保存9张!');}</script>"
		Response.End
end if
if rs("i")>=9 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你的像册已经多于9张，请删除再保存!');}</script>"
		Response.End
end if
if instr(rs("g"),data)<>0 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：此张像册你已经保存过!');}</script>"
		Response.End
end if
conn.execute("update use set f='"& data &"' where a='" & aqjh_name & "'")
if trim(Request("act"))="savephoto" then
	conn.execute("update use set i=i+1, g=g & '"& data &";' where a='" & aqjh_name & "'")
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：保存完成!');}</script>"
Response.End
%>