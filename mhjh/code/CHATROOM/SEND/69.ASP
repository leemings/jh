<%
function jijin(un,co)
    if co="无" then
     Response.Write "<script language=JavaScript>{alert('"&un&",你还没有加任何门派!');}</script>"
	 Response.End
	 exit function
	 end if
rst.Open "select 等级,精力,银两 from 用户 where 姓名='"&un&"'",conn
if rst("等级")>3 and rst("精力")>1000 and rst("银两")>1000 then
conn.execute "update 用户 set 精力=精力-1000,银两=银两-1000 where 姓名='"&un&"'" 
rst.Close
rst.Open "select 基金 from 门派 where 门派='"&co&"'",conn
jj=rst("基金")
if jj<1000000 then
jijin="<marquee border='1' scrollamount=3><font color=FF0000><font face=Webdings class=p>&yuml;</font>【基金状况】</font>##,你门派基金还有"&jj&"两,请及时补充本门后备资源!</marquee>"
else
jijin="<marquee border='1' scrollamount=3><font color=FF0000><font face=Webdings class=p>&yuml;</font>【基金状况】</font>##,你门派基金还有"&jj&"两,还比较充足的!</marquee>"
end if
else
response.write "<script language='javascript'>alert('你连走路都没劲还想查看本门基金');</script>"
response.end
end if
rst.Close
end function
%>