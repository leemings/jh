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
id=Request("id")
sl=Request("sl")
sl=1
if (not isnumeric(id)) or (not isnumeric(sl))  then 
	Response.Write "<script language=JavaScript>{alert('提示：数据提交有误！');window.close();}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from use where a='"& aqjh_name &"'", conn, 1,3
if rs.eof and rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "reg.asp"
	Response.End
end if
if rs("e")<=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：金币数量为0，请去充值！');window.close();}</script>"
	Response.End
end if
mysex=rs("c")'性别
myhy=rs("d")'会员
myjb=rs("e")'金币
rs.close
rs.Open "select * from data where id="& id, conn, 1,3
if rs.eof and rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：所购买商品并不存在！');window.close();}</script>"
	Response.End
end if
rs("l")=rs("l")+1
rs.update
spname=rs("b")
spnj=rs("h")
'非会员判断
if rs("e")=true and myhy=false then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不是会员不能购买会员商品！');window.close();}</script>"
	Response.End
end if
'会员商品
if myhy=false then
	spjb=rs("f")
else
	spjb=rs("g")
end if
if myjb<spjb then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币数量不足！');window.close();}</script>"
	Response.End
end if
if rs("d")<>mysex and rs("d")<>"U" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：性别不符，不可以购买！');window.close();}</script>"
	Response.End
end if
rs.close
dim tmprs
tmprs=conn.execute("Select count(*) As 数量 from usesp where sa='"& sjjhname &"'")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
if myhy=false and musers>=10 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：非会员限制只可以购买10种商品！');window.close();}</script>"
	Response.End
end if
if myhy=true and musers>=30 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：会员限制可以购买30种商品！');window.close();}</script>"
	Response.End
end if
tmprs=conn.execute("Select count(*) As 数量 from usesp where sb="& id)
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
if musers<>0 and spjb=0  then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：免费商品只可以购买1个！');window.close();}</script>"
	Response.End
end if
rs.Open "select top 1 * from usesp where DateDiff('d',date(),sd)<=0", conn, 1,3
if rs.eof and rs.bof then
rs.addnew
end if
rs("sa")=aqjh_name
rs("sb")=id
rs("sc")=date()
rs("sd")=DateAdd("d",spnj,date())
rs("se")=false
rs.update
rs.close
if spjb>=0 then
	rs.Open "select * from use where a='"& aqjh_name &"'", conn, 1,3
	rs("e")=rs("e")-spjb
	rs.update
	rs.close
end if
set rs=nothing
conn.close
set conn=nothing
if spjb<=0 then
	Response.Write "<script language=JavaScript>{alert('提示：["& spname &"]购买1个成功，总数:"& musers+1 &"个!');window.close();}</script>"
else
	Response.Write "<script language=JavaScript>{alert('提示：["& spname &"]购买1个成功，总数:"& musers+1 &"个,花费("& spjb &")个金币!');window.close();}</script>"
end if
Response.End
%>