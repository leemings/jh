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
hygrade=int(Request.form("hygrade"))
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
	'if rs("��Ա�ȼ�")<>0 or rs("grade")=10 then
	'	Response.Write "<script language=javascript>alert('��Ǹ["&username&"]�����Ѿ���["&rs("��Ա�ȼ�")&"��]������Ա�����ֶ����ģ�');history.back();</script>"
	'	rs.close
	'	set rs=nothing
	'	conn.close
	'	set conn=nothing
	'	Response.End	
	'end if
	if rs("��Ա�ȼ�")>1 then
		hydate=rs("��Ա����")
	else
		hydate=date()
	end if
	oldhy=rs("��Ա�ȼ�")
	n=Year(hydate)
	y=Month(hydate)
	r=Day(hydate)
	y=y+hymonth
	if y>12 then 
		y=y-12
		n=n+1
	end if
	jy=0
	if hygrade=1 then jy=31250
	if hygrade=2 then jy=90000
	if hygrade=3 then jy=250000
	if hygrade=4 then jy=490000
	if hygrade=5 then jy=1000000
	if hygrade=6 then jy=1800000
	if hygrade=7 then jy=2000000
	if hygrade=8 then jy=4500000
	if oldhy=1 then oldjy=31250
	if oldhy=2 then oldjy=90000
	if oldhy=3 then oldjy=250000
	if oldhy=4 then oldjy=490000
	if oldhy=5 then oldjy=1200000
	if oldhy=6 then oldjy=1800000
	if oldhy=7 then oldjy=2000000
	if oldhy=8 then oldjy=4500000
	sj=n & "-" & y & "-" & r
	if hygrade=8 then
                if username=Application("aqjh_user") then
                   response.write "<script language=javascript>alert('������Ķ����ǽ���10������Ա,�����ô˲�����');history.back();</script>"
                   response.end
                end if
		rs("����")="�ٸ�"
		rs("grade")=7
	end if
	rs("��Ա�ȼ�")=hygrade
	rs("��Ա����")=cdate(sj)
	rs("allvalue")=rs("allvalue")+jy-oldjy
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','[���ӵȼ��ƻ�Ա]������"&username&"���ȼ���"&hygrade&" ���ڣ�"&sj&"')"
	set rs=nothing
	conn.Close
	set conn=nothing
	Response.Write "<script language=javascript>alert('��ϲ��["&username&"]�Ľ���["& hygrade &"]��Ա������ɣ�');history.back();</script>"
%>