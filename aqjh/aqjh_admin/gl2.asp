<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="delete * from �û� where  allvalue<=5 and CDate(lasttime)<now()-10"
Set Rs=conn.Execute(sql)
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','10��֮ǰ�ݵ���allvalueС��5���û��Ѿ�ȫ��ɾ����')"
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('10��֮ǰ�ݵ���allvalueС��5���û��Ѿ�ȫ��ɾ����');"
Response.Write "location.href = 'manacc.asp';"
Response.Write "</script>"
Response.End
%>