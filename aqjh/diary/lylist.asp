<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
id=request("id")
if id="" or (not isnumeric(id)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ڸ�ʲô��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From ����  where ����id="& id &" order by ʱ�� DESC", conn1, 1,1
if rs1.bof and rs1.eof then
	response.write "document.write('<p align=center><font color=red>��ʱû�����ԡ�</font></p>');"
else
news=1
do while not rs1.eof and news<=5
lyid=rs1("id")
name=rs1("����")
lysj=rs1("ʱ��")
if news/2=int(news/2) then
	response.write "document.write("& chr(34) &"<div style=background:#2DBBFF>"&news&":<a href=# onClick=window.open('lyshow.asp?id="& rs1("id") &"','seely','scrollbars=yes,resizable=no,width=280,height=300') title='"& lysj &"'>"& badstr(rs1("����")) &" By "& name &"</a>&nbsp;&nbsp;<a href=diary.asp?action=lydel&id="& id &"&lyid="& lyid &"><img src=images/delete.gif width=13 height=13 border=0 alt=ɾ�������ԣ�></a></div>"& chr(34) &");"
else
	response.write "document.write("& chr(34) &"<div style=background:#66CCFF>"&news&":<a href=# onClick=window.open('lyshow.asp?id="& rs1("id") &"','seely','scrollbars=yes,resizable=no,width=280,height=300') title='"& lysj &"'>"& badstr(rs1("����")) &" By "& name &"</a>&nbsp;&nbsp;<a href=diary.asp?action=lydel&id="& id &"&lyid="& lyid &"><img src=images/delete.gif width=13 height=13 border=0 alt=ɾ�������ԣ�></a></div>"& chr(34) &");"	
end if
news=news+1
rs1.movenext
loop
end if
rs1.close
set rs=nothing
conn1.close
set conn1=nothing 
Response.Write "document.write('<div align=right><a href="& chr(34) &"more.asp?id="& id & chr(34) &" target=blank><img border=0 src="& chr(34) &"images/gd.gif"& chr(34) &" alt="& chr(34) &"��������"& chr(34) &">�鿴���੬</a></div>');"
function badstr(str)
	badword="�侫,��,ʺ,��,��,��,��,����,��,��,��,�,����,����,����,����,غ��,��,��Ƥ,��ͷ,��,�P,��,�H,��,��,��,����,����,��,����,����,��Ů,����,ʮ����,��,�Ҷ�,��,��ϯ,����,����,��־,��,����,����,����,��Ժ,����,����,��"
	bad=split(badword,",")
	for i=0 to ubound(bad)-1
		if InStr(LCase(str),bad(i))<>0 then 
			str=replace(str,bad(i),"*")
		end if
	next
	badstr=str
end function
%>