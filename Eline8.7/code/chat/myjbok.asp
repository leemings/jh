<%@ LANGUAGE=VBScript codepage ="936" %> 
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
wpname=Trim(Request.Form("wpname"))
money=clng(Trim(Request.Form("money")))
ggsm=Trim(Request.Form("ggsm"))
if not isnumeric(money) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ʹ�����֣�');}</script>"
	Response.End 
end if
if len(ggsm)>150 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Գ���150������!');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b where a='"&wpname&"' and (b='ҩƷ' or b='�ʻ�')  order  by b,o",conn,2,2
if (rs("n")="" or isnull(rs("n")) or rs("n")<>sjjh_name or rs("n")="��") and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&wpname&"]Ҳ�����㾭Ӫ��!');window.close();}</script>"
	Response.End 
end if
if rs("b")="�ʻ�" then
	miniwj=50000
	maxwj=2000000
else
	miniwj=(rs("d")+rs("e"))*20
	maxwj=(rs("d")+rs("e"))*50
end if
if money<miniwj or money>maxwj then
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&wpname&"]��ͼ�Ϊ:"&miniwj&"������߼�Ϊ:"&maxwj&"��!');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
rs.close
sqlstr="SELECT * FROM b where a='"&wpname&"'"
rs.Open sqlstr,conn,1,2
rs("h")=money
rs("p")=ggsm
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ������Ӫ����Ʒ["&wpname&"]�޸ĳɹ�!');location.href = 'myjb.asp?wpname="&wpname&"';</script>"
%>
