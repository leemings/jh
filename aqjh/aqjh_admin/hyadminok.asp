<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
username=Request.Form("search")
sjyear=Request.form("year")
sjmonth=Request.form("month")
sjday=Request.form("day")
hygrade=int(Request.form("hygrade"))
hydate="#"&sjyear&"/"&sjmonth&"/"&sjday&"#"
select case hygrade
case 1
jhgrade=1
hyvalue=31250
hyjia=0
case 2
jhgrade=1
hyvalue=90000
hyjia=100
case 3
jhgrade=5
hyvalue=250000
hyjia=200
case 4
jhgrade=1
hyvalue=490000
hyjia=300
case 5
jhgrade=1
hyvalue=1200000
hyjia=400
case 6
jhgrade=5
hyvalue=1800000
hyjia=500
case 7
jhgrade=6
hyvalue=2000000
hyjia=600
case 8
jhgrade=7
hyvalue=4500000
hyjia=700
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
Set rs=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM �û� where ����='"&username&"'"
rs.Open sqlstr,conn
if rs.EOF or rs.BOF then
	Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ��������Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
else
	if rs("��Ա�ȼ�")=0 and rs("grade")<>10 then
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','[���ӻ�Ա]������"&username&"���ȼ���"&hygrade&" ���ڣ�"&hydate&"')"
		sql="update �û� set ��Ա�ȼ�="& hygrade &",��Ա����="& hydate &",������=������+"& hyjia &",������=������+"& hyjia &",�书��=�书��+"& hyjia &",grade="& jhgrade &",allvalue=allvalue+"& hyvalue &" where ����='"&username&"'"
		Set rs=conn.Execute(sql)
	else
		Response.Write "<script language=javascript>alert('��Ǹ["&username&"]�����Ѿ���["&rs("��Ա�ȼ�")&"��]������Ա�����ֶ����ģ�');history.back();</script>"
	end if



Response.Write "<script language=javascript>alert('��ϲ��["&username&"]�Ľ���["& hygrade &"]��Ա������ɣ�');history.back();</script>"
%>