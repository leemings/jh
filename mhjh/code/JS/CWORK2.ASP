<title>职业转换</title>
<body bgcolor="#005b00">
<%
my=session("yx8_mhjh_username")
if my=""  then%> 
<script language="vbscript">
  alert("你连接超时了！")
</script>
<%
response.end
end if
jiu=request.form("jiu")
	select case jiu
           case "cb"
              ziye="冰工"
           case "ck"
              ziye="矿工"
           case "du"
              ziye="渔夫"
           case "mao"
              ziye="探险"
           case "yao"
              ziye="郎中"
           case "chu"
              ziye="厨师"
          end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
	<script language="vbscript">
		alert("您不是江湖中人，我们不欢迎你！")
	  	window.close()
	</script>
<%
conn.close
response.end
end if
if rs("银两")<100000 then
%>
	<script language="vbscript">
         	alert("如果您的钱不够10万，这辈子做牛做马你也还不起,穷鬼走开些!!")
               window.close()
</script>
<%
end if
if rs("银两")>100000 then
sql="update 用户 set 银两=银两-100000,职业='"& ziye &"' where 姓名='" & my & "'"
conn.execute sql
conn.close
set rs=nothing
Response.Write "<script Language=Javascript>alert('您职业转换成了："& ziye &"工作，点击确定返回！');</script>"
end if
%>
<script language=javascript>
location.href='cwork1.asp'
</script>

