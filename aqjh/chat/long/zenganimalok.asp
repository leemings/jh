<%
if session("aqjh_grade")<10 then Response.Redirect "../../error.asp?id=440"
name=trim(Request.Form ("zengname"))
animal=trim(Request.Form ("animal"))
attack=trim(Request.Form ("attack"))
if not isnumeric(attack) then 
Response.Write "�Բ���������Ĺ�������Ϊ���֣�"
Response.End 
end if
if InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name,"��")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 then Response.Redirect "error.asp?id=470"
if InStr(animal,"=")<>0 or InStr(animal,"`")<>0 or InStr(animal,"'")<>0 or InStr(animal," ")<>0 or InStr(animal,"��")<>0 or InStr(animal,"'")<>0 or InStr(animal,chr(34))<>0 or InStr(animal,"\")<>0 or InStr(animal,",")<>0 or InStr(animal,"<")<>0 or InStr(animal,">")<>0 then Response.Redirect "error.asp?id=470"
if attack>1000000 then attack=100000 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("select id from �û� where ����='"&name&"'"),conn
if rs.BOF or rs.EOF then 
Response.Write "�Բ���["&name&"]û�з��ִ��˵����ϣ�"
else
conn.execute("insert into myanimal(animalname,username,attack,outtime,rest) values('"&animal&"','"&name&"',"&attack&",0,0)")
end if
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
Response.Write "���ͳɹ���<a href=javascript:parent.history.go(-1);>���˴�����</a>"
%>
</body>
</html>
