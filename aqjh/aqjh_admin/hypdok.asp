<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
username=Request.Form("search")
hymonth=Request.form("hymonth")
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	sqlstr="SELECT * FROM �û� where ����='"&username&"'"
	rs.Open sqlstr,conn,1,2
	if rs.EOF or rs.BOF then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ��������Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
		Response.End
	end if
	if rs("��Ա")=true then
		hydate=rs("��Ա����")
	else
		hydate=date()
	end if
	n=Year(hydate)
	y=Month(hydate)
	r=Day(hydate)
	y=y+hymonth
	if y>12 then 
		y=y-12
		n=n+1
	end if
	sj=n & "-" & y & "-" & r
	rs("��Ա")=true
	rs("��Ա����")=cdate(sj)
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','[�����ݵ��ƻ�Ա]������"&username&"�����ڣ�"&sj&"')"
	set rs=nothing
	conn.Close
	set conn=nothing
	Response.Write "<script language=javascript>alert('��ϲ��["&username&"]�Ľ���[�ݵ���]��Ա������ɣ�');history.back();</script>"
%>