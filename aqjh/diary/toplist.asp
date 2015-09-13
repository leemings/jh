<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
	if trim(request("n"))<>"" and IsNumeric(request("n")) then
	n=cint(request("n"))
	else
	n=1
	end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From 日记 Order By 书写日期 DESC", conn1, 1,1
if rs1.bof and rs1.eof then
	response.write "document.write('<p align=center><font color=red>暂时没有日记。</font></p>');"
else
news=1
do while not rs1.eof and news<=n
select case rs1("保密")
case "1"
	bm="公开欣赏"
case "2"
	bm="保密内容!"
case "3"
	bm="凭密码查看"
case "4"
	bm="[" & rs1("保密条件") & "]查看"
case "5"
	bm="[" & rs1("保密条件") & "]级以上"
case "6"
	bm="[" & rs1("保密条件") & "]查看"
end select
titlename=rs1("日记标题")
		if len(titlename)>Cint(request("tlen")) then
			titlename=left(titlename,request("tlen"))&"..."
		end if
response.write "document.write('<img height=14 src=diary/images/hot.gif width=11>   <a href=diary/show.asp?id=" & rs1("id") & " & wj=list.asp title=" & chr(34) & "作者：" & rs1("用户名") & " 鲜花:" & rs1("鲜花") & " 鸡蛋:" & rs1("鸡蛋") & " 留言:" & rs1("留言数") & " 方式:" & bm & "" & chr(34) & " target=blank>" & titlename & "</a><br>');"
news=news+1
rs1.movenext
loop
end if
rs1.close
set rs=nothing
conn1.close
set conn1=nothing 
Response.Write "document.write('<div align=right><a href=" & chr(34) & "diary/list.asp " & chr(34) & "target=blank><img border=0 src=" & chr(34) & "diary/images/gd.gif" & chr(34) & " alt=" & chr(34) & "查看更多日记" & chr(34) & ">   查看更多┈</a></div>');"
function badstr(str)
	badword="射精,奸,屎,妈,娘,尻,操,王八,逼,贱,狗,婊,叉你,插你,干你,鸡巴,睾丸,蛋,包皮,龟头,屄,赑,妣,肏,奶,尻,屌,作爱,做爱,床,抱抱,鸡八,处女,打炮,十八摸,爸,我儿,·,主席,泽民,法伦,洪志,大法,公安,政府,反动,法院,升天,周天,功,'"
	bad=split(badword,",")
	for i=0 to ubound(bad)-1
		if InStr(LCase(str),bad(i))<>0 then 
			str=replace(str,bad(i),"*")
		end if
	next
	badstr=str
end function
%>