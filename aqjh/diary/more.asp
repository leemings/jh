<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
function badstr(str)
	badword="�侫,��,ʺ,��,��,��,��,����,��,��,��,�,����,����,����,����,غ��,��,��Ƥ,��ͷ,��,�P,��,�H,��,��,��,����,����,��,����,����,��Ů,����,ʮ����,��,�Ҷ�,��,��ϯ,����,����,��־,��,����,����,����,��Ժ,����,����,��"
	bad=split(badword,",")
	for ti=0 to ubound(bad)-1
		if InStr(LCase(str),bad(ti))<>0 then 
			str=replace(str,bad(ti),"*")
		end if
	next
	badstr=str
end function
%>
<html>
<title>�鿴����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/diary.css" rel=stylesheet type=text/css>
<body background=images/bg1.gif>
<%
id=request("id")
if id="" or (not isnumeric(id)) then
Response.Write "<script language=JavaScript>{alert('��ʾ�����ڸ�ʲô��');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From ����  where ����id="& id &" order by ʱ�� DESC", conn1, 1,1
if rs1.eof and rs1.bof then%>
��ʱû���κμ�¼����
<%else
RS1.PageSize=5'����ÿҳ��¼��,���޸�
Dim TotalPages
TotalPages = RS1.PageCount
If CurPage>RS1.Pagecount Then
CurPage=RS1.Pagecount
end if
RS1.AbsolutePage=CurPage
rs1.CacheSize = RS1.PageSize'��������¼��
Dim Totalcount
Totalcount =INT(RS1.recordcount)
StartPageNum=1
do while StartPageNum+10<=CurPage
StartPageNum=StartPageNum+10
Loop
EndPageNum=StartPageNum+9
If EndPageNum>RS1.Pagecount then EndPageNum=RS1.Pagecount%>
<br>
<table border="1" width="70%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
<tr>
<td bordercolorlight="#00215a" align=center bgcolor="#d7ebff"><B>��ѯ����</B>
</td>
</tr>
<%I=0
p=RS1.PageSize*(Curpage-1)
do while (Not RS1.Eof) and (I=<RS1.PageSize)
p=p+1
lyid=rs1("id")
name=rs1("����")
lysj=rs1("ʱ��")
if i/2=int(i/2) then
	response.write "<tr><td ><div style=background:#2DBBFF>����"& i &":<a href=# onClick=window.open('lyshow.asp?id="& lyid &"','seely','scrollbars=yes,resizable=no,width=280,height=300') title='"& lysj &"'>"& badstr(rs1("����")) &"</a>&nbsp;&nbsp; By "& name &"&nbsp;&nbsp; <a href=diary.asp?action=lydel&id="& id &"&lyid="& lyid &"><img src=images/delete.gif width=13 height=13 border=0 alt=ɾ�������ԣ�></a></div></td></tr>"
else
	response.write "<tr><td ><div style=background:#66CCFF>����"&i&":<a href=# onClick=window.open('lyshow.asp?id="& lyid &"','seely','scrollbars=yes,resizable=no,width=280,height=300') title='"& lysj &"'>"& badstr(rs1("����")) &"</a>&nbsp;&nbsp; By "& name &"&nbsp;&nbsp; <a href=diary.asp?action=lydel&id="& id &"&lyid="& lyid &"><img src=images/delete.gif width=13 height=13 border=0 alt=ɾ�������ԣ�></a></div></td></tr>"	
end if
%>
<%I=I+1
RS1.MoveNext
Loop%>
<tr>
<td colspan=4 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="50%">
</td>
<td width="50%" align="right"> ��<%=Totalcount%>������ [
<% if StartPageNum>1 then %>
<a href="more.asp?CurPage=<%=StartPageNum-1%>"><<</a>
<%end if%>
<% For I=StartPageNum to EndPageNum
if I<>CurPage then %>
<a href="more.asp?CurPage=<%=I%>"><%=I%></a>
<% else %>
<b><font color=#ff0000><%=I%></font></b>
<% end if %>
<% Next %>
<% if EndPageNum<RS1.Pagecount then %>
<a href="more.asp?CurPage=<%=EndPageNum+1%>">>></a>
<%end if%>
] </td>
</tr>
</table>
</td>
</tr>
</table>
<p align=center>���齭��</p></body></html>
<%
end if
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
%>