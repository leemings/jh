<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
<%
Response.Buffer=true
username=session("aqjh_name")
if username="" then Response.Redirect "../../error.asp?id=440"
mj=request.form("mj")
if not isnumeric(mj) then Response.Redirect  "../../error.asp?id=464"
if mj="" or mj<10000 then
Response.Write username&"��̫С���˰ɣ�10000������Ҳ�費����"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
const MaxPerPage=10
sql="select * from �û�  where ����='"&username&"'"
set rs=conn.execute(sql)
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');window.close()}</script>"
	Response.End
end if
if rs("����")-mj<0 then
Response.Write username&"������Ǯ�����ɣ�ûǮ��ʲô���?"
Response.End 
end if
meili=int(mj*.0002)
daode=int(mj*.0003)
fang=int(mj*.00005)
gong=int(mj*.00005)
sql="update �û� set ����=����-"&mj&",����=����+"&meili&",����=����+"&daode&",����=����+"&fang&",����=����+"&gong&",����ʱ��=now() where ����='"&username&"'"
Response.Write sql
set rs=conn.Execute(sql)
conn.Close 
set conn=nothing
Response.Redirect "mujuan.asp"
%>
<head>
<title></title>
</head><body bgcolor="#000000" text="#FFFFFF">