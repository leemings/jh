<%
function bujin(un,co,mg)
    if not isnumeric(mg) then
     Response.Write "<script language=JavaScript>{alert('"&un&",请输入数字!');}</script>"
	 Response.End
	 exit function
	 end if
	 if mg<=0 then
      Response.Write "<script language=JavaScript>{alert('"&un&",请输入大于0的数字!');}</script>"
	 Response.End
	 exit function
	 end if
rst.Open "select 等级 from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 银两>"&mg&"",conn
if rst.EOF or rst.BOF then
bujin="<font color=FF0000>【门派基金入库】</font>##,你不是本派的，或者你的银两输入比你自己的还多"
else
conn.execute "update 用户 set 银两=银两-"&mg&" where  姓名='"&un&"'"
conn.execute "update 门派 set 基金=基金+"&mg&" where  门派='"&co&"'"
bujin="<font color=FF0000>【门派基金入库】</font>##,你为了本门继续发扬光大,不惜把"&mg&"贡献给本门金库!"
end if
rst.Close
end function
%>
