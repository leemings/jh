<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
id=request("id")
if id="" or (not isnumeric(id)) then 
	Response.Write "<script language=JavaScript>{alert('提示：你在搞什么！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From 留言  where 留言id="& id &" order by 时间 DESC", conn1, 1,1
if rs1.bof and rs1.eof then
	response.write "document.write('<p align=center><font color=red>暂时没有留言。</font></p>');"
else
news=1
do while not rs1.eof and news<=5
lyid=rs1("id")
name=rs1("姓名")
lysj=rs1("时间")
if news/2=int(news/2) then
	response.write "document.write("& chr(34) &"<div style=background:#2DBBFF>"&news&":<a href=# onClick=window.open('lyshow.asp?id="& rs1("id") &"','seely','scrollbars=yes,resizable=no,width=280,height=300') title='"& lysj &"'>"& badstr(rs1("标题")) &" By "& name &"</a>&nbsp;&nbsp;<a href=diary.asp?action=lydel&id="& id &"&lyid="& lyid &"><img src=images/delete.gif width=13 height=13 border=0 alt=删除此留言！></a></div>"& chr(34) &");"
else
	response.write "document.write("& chr(34) &"<div style=background:#66CCFF>"&news&":<a href=# onClick=window.open('lyshow.asp?id="& rs1("id") &"','seely','scrollbars=yes,resizable=no,width=280,height=300') title='"& lysj &"'>"& badstr(rs1("标题")) &" By "& name &"</a>&nbsp;&nbsp;<a href=diary.asp?action=lydel&id="& id &"&lyid="& lyid &"><img src=images/delete.gif width=13 height=13 border=0 alt=删除此留言！></a></div>"& chr(34) &");"	
end if
news=news+1
rs1.movenext
loop
end if
rs1.close
set rs=nothing
conn1.close
set conn1=nothing 
Response.Write "document.write('<div align=right><a href="& chr(34) &"more.asp?id="& id & chr(34) &" target=blank><img border=0 src="& chr(34) &"images/gd.gif"& chr(34) &" alt="& chr(34) &"更多留言"& chr(34) &">查看更多┈</a></div>');"
function badstr(str)
	badword="射精,奸,屎,妈,娘,尻,操,王八,逼,贱,狗,婊,叉你,插你,干你,鸡巴,睾丸,蛋,包皮,龟头,屄,赑,妣,肏,奶,尻,屌,作爱,做爱,床,抱抱,鸡八,处女,打炮,十八摸,爸,我儿,·,主席,泽民,法伦,洪志,大法,公安,政府,反动,法院,升天,周天,功"
	bad=split(badword,",")
	for i=0 to ubound(bad)-1
		if InStr(LCase(str),bad(i))<>0 then 
			str=replace(str,bad(i),"*")
		end if
	next
	badstr=str
end function
%>