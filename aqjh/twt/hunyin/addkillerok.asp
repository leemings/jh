<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
name=trim(Request.form("name"))
mess=trim(Request.form("mess"))
mess=LCase(trim(Request.form("mess")))
mess=replace(mess,"'","")
mess=replace(mess,chr(34),"")
mess=Replace(mess,"<","")
mess=Replace(mess,">","")
mess=Replace(mess,"\x3c","")
mess=Replace(mess,"\x3e","")
mess=Replace(mess,"\074","")
mess=Replace(mess,"\74","")
mess=Replace(mess,"\75","")
mess=Replace(mess,"\76","")
mess=Replace(mess,"&lt","")
mess=Replace(mess,"&gt","")
mess=Replace(mess,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
mess=Replace(mess,bad(i),"**")
next
money=int(abs(Request.form("money")))
if name=aqjh_name or name="" or mess="" then 
	Response.Write "<script Language=Javascript>alert('��ʾ����������Ϊ�գ�����ɱ�Լ�������');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if money<10000 or money>100000000 then
	Response.Write "<script Language=Javascript>alert('��ʾ��������ô���Ǯ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM h WHERE a='"&aqjh_name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����Ҫɱ���˲����ڡ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
'У���û� ��������100��Ǯ����1000
rs.open "SELECT * FROM �û� WHERE ����='"&aqjh_name&"'"&" and ����>"&money,conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��������ô���Ǯ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT * FROM e WHERE a='"&aqjh_name&"' and b='"&name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ�������˲����ظ�������');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
	conn.execute "update �û� set ����=����-"&money&" where ����='"&aqjh_name&"'"
	conn.Execute "INSERT INTO e (a,b,c,d,e,f) VALUES('"&aqjh_name&"','"&name&"','"&mess&"',now(),"&money&",'��')"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�������ύ�ɹ���');location.href='killer.asp';</script>"
response.end
%>