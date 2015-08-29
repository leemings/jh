<%
Response.Buffer=true
username=session("sjjh_name")
if username="" then Response.Redirect "error.asp?id=440"
mj=request.form("mj")
if not isnumeric(mj) then Response.Redirect  "error.asp?id=464"
if mj="" or mj<10000 then
Response.Write username&"你太小气了吧，10000两银子也捐不出吗？"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
const MaxPerPage=10
sql="select * from 用户  where 姓名='"&username&"'"
set rs=conn.execute(sql)
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"分钟]再操作！');}</script>"
	Response.End
end if

if rs("银两")-mj<0 then
Response.Write username&"您好像钱不够吧，没钱充什么大款?"
Response.End 
end if
meili=int(mj*.0002)
daode=int(mj*.0003)
fang=int(mj*.00005)
gong=int(mj*.00005)
sql="update 用户 set 银两=银两-"&mj&",魅力=魅力+"&meili&",道德=道德+"&daode&",防御=防御+"&fang&",攻击=攻击+"&gong&",操作时间=now() where 姓名='"&username&"'"
Response.Write sql
set rs=conn.Execute(sql)
conn.Close 
set conn=nothing
Response.Redirect "mujuan.asp"
%>

<head>
<title>♀wWw.51eline.com♀</title>
</head>

<body bgcolor="#000000" text="#FFFFFF">