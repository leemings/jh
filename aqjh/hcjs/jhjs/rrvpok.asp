<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
'http = Request.ServerVariables("HTTP_REFERER") 
'if InStr(http,"hcjs/jhjs")=0 then 
'	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
'	Response.End 
'end if
vhtype=cint(trim(Request.Form("type")))
vhid=trim(Request.Form("id"))
if not isnumeric(vhid) then
	Response.Write "<script language=javascript>{alert('���ɣ�������ʲô�����뵷��ѽ������');window.close();}</script>" 
	Response.End 
end if
if vhtype=2 then
	tt=clng(trim(Request.Form("tt")))
	vhrnd=clng(trim(Request.Form("vhrnd")))/100
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT a,k FROM vh where id="&vhid,conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End
end if
if rs("k")>4 then
	conn.execute "delete * from vh where id="&vhid
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ݸñ����ˣ����ٸ�ǿ�Ʊ��ϣ�');parent.location='vehicle.asp';}</script>"
	Response.End
end if
vhname=rs("a")
rs.close
rs.open "select h,j,m from b where a='"&vhname&"'",conn
randomize()
if vhtype=1 then
	vhh=int(rs("h")*0.3)
	vhm=int(rs("m")*0.6+0.5)
	vhj=int(rs("j")*(rnd()*0.1+0.9))
else
	vhtt=int(rs("m")/vhrnd+sqr(rs("h"))*vhrnd)
	if vhtt<>tt then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>{alert('���ɣ�������ʲô����������ѽ������');window.close();}</script>" 
		Response.End 
end if
	vhh=int(rs("h")*vhrnd+0.5)
	vhm=int(rs("m")*vhrnd+0.5)
	vhj=int(rs("j")*(rnd()*0.24+2*vhrnd))
if rnd()<0.2 then vhj=0
end if
rs.close
rs.open "select ����,��� from �û� where ����='"&aqjh_name&"'",conn
conn.execute "update �û� set ����=����+"&vhh&",���=���+"&vhm&" where ����='"&aqjh_name&"'"
conn.execute "delete * from vh where id="&vhid
rs.close
set rs=nothing
conn.close
set conn=nothing
if vhj=0 then
	Response.Write "<script Language=Javascript>{alert('Ӣ��ûǮ,�޵�֮��,Ϊ��������Ѱ�������,���˵õ�����"&vhh&"�������"&vhm&"��!');parent.location='vehicle.asp';}</script>"
else
	Response.Write "<script Language=Javascript>{alert('Ӣ��ûǮ,�޵�֮��,Ϊ��������Ѱ�������,���˵õ�����"&vhh&"�������"&vhm&"��!');parent.location='vehicle.asp';}</script>"

end if
%>