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
rs1.Open "Select * From �ռ� Order By ��д���� DESC", conn1, 1,1
if rs1.bof and rs1.eof then
	response.write "document.write('<p align=center><font color=red>��ʱû���ռǡ�</font></p>');"
else
news=1
do while not rs1.eof and news<=n
select case rs1("����")
case "1"
	bm="��������"
case "2"
	bm="��������!"
case "3"
	bm="ƾ����鿴"
case "4"
	bm="[" & rs1("��������") & "]�鿴"
case "5"
	bm="[" & rs1("��������") & "]������"
case "6"
	bm="[" & rs1("��������") & "]�鿴"
end select
titlename=rs1("�ռǱ���")
		if len(titlename)>Cint(request("tlen")) then
			titlename=left(titlename,request("tlen"))&"..."
		end if
response.write "document.write('<img height=14 src=diary/images/hot.gif width=11><a href=diary/show.asp?id=" & rs1("id") & " & wj=list.asp title=" & chr(34) & "���ߣ�" & rs1("�û���") & " �ʻ�:" & rs1("�ʻ�") & " ����:" & rs1("����") & " ����:" & rs1("������") & " ��ʽ:" & bm & "" & chr(34) & " target=blank>" & titlename & "</a><br>');"
news=news+1
rs1.movenext
loop
end if
rs1.close
set rs=nothing
conn1.close
set conn1=nothing 
Response.Write "document.write('<div align=right><a href=" & chr(34) & "diary/list.asp " & chr(34) & "target=blank><img border=0 src=" & chr(34) & "diary/images/gd.gif" & chr(34) & " alt=" & chr(34) & "�鿴�����ռ�" & chr(34) & ">�鿴���੬</a></div>');"
function badstr(str)
	badword="�侫,��,ʺ,��,��,��,��,����,��,��,��,�,����,����,����,����,غ��,��,��Ƥ,��ͷ,��,�P,��,�H,��,��,��,����,����,��,����,����,��Ů,����,ʮ����,��,�Ҷ�,��,��ϯ,����,����,��־,��,����,����,����,��Ժ,����,����,��,'"
	bad=split(badword,",")
	for i=0 to ubound(bad)-1
		if InStr(LCase(str),bad(i))<>0 then 
			str=replace(str,bad(i),"*")
		end if
	next
	badstr=str
end function
%>