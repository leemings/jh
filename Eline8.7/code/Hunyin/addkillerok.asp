<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
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
badstr="������վ|���ͽ���|�����콭��|��������|����|����|�ĸ�|����|����|tama|ta ma|tamu|ta mu|ѧ�����|�ƽ���|���Ľ���|ȥ��Ľ���|ȥ�ҵĽ���|���ҵĽ���|���ҵĽ���|ȥ��������|ȥ��������|����������|����������|�ý���|�Ľ���|�侫|��|�Ҹ�|�ٸ�|����|��ȫ|��ʲ|��ʲ|cfq|���������|���������|���������|������|chenfaquan|chen fa quan|chen|ma de|kan ni|fa|quan|ni ma|mao|ni mu|niba|ni ba|nima|nimu|ni nai nai|ni jie|ni mei|nai|ye|ma|ba|zu|zhu|ɵ��|վ��|ռ��|ս��|�ɳ�|�ʳ�|����|�ҿ�|��ʦ|����|������|������|1��|ͱ��|ͱ����|��ʦ|��Ů|�·�|��Ѽ|��Ѽ|�ϻ�|sai|����Ա|����|kao|cao|yixiantian|yi xian tian|yxt|eline|����ĸ|8196316|13808556346|һ��|һ����|����|һǮ��|��ĸ|��ľ|����|ȥ��|վ��|�ϴ�|��ʺ|��|��|��|��|����|��|үү|����|��ү|����|��ү|����|����|�Ҳ�|������|����|��|ɵ|��|������|���|����|���|����|�|����|����|��|����|����|����|����|����|����|غ��|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|����|��|�Ҷ�|���|���|��|����|��ð|��ï|��ð|��ï|����|�ڹ�|ʺ|ƨ|����|�׳�|��|������|���|���˹�|http|www|com|cn|org|net|//|fuck|cao|gan|sai|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
mess=Replace(mess,bad(i),"**")
next
money=int(abs(Request.form("money")))
if name=sjjh_name or name="" or mess="" then 
	Response.Write "<script Language=Javascript>alert('��ʾ����������Ϊ�գ�����ɱ�Լ�������');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if money<10000 or money>100000000 then
	Response.Write "<script Language=Javascript>alert('��ʾ��������ô���Ǯ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM h WHERE a='"&sjjh_name&"'",conn
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
rs.open "SELECT * FROM �û� WHERE ����='"&sjjh_name&"'"&" and ����>"&money,conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��������ô���Ǯ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT * FROM e WHERE a='"&sjjh_name&"' and b='"&name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ�������˲����ظ�������');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
	conn.execute "update �û� set ����=����-"&money&" where ����='"&sjjh_name&"'"
	conn.Execute "INSERT INTO e (a,b,c,d,e,f) VALUES('"&sjjh_name&"','"&name&"','"&mess&"',now(),"&money&",'��')"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�������ύ�ɹ���');location.href = 'killer.asp';</script>"
response.end
%>