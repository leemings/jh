<%
function chuanyan(un,st)
rst.Open "select 等级,精力,银两 from 用户 where 姓名='"&un&"'",conn
if rst("等级")>6 and rst("精力")>1000 and rst("银两")>100000 then
conn.execute "update 用户 set 精力=精力-100,银两=银两-100000 where 姓名='"&un&"'" 
chuanyan="<marquee border='1'><font color=FF0000><font face=Webdings class=p>&yuml;</font>【千里传音】（精力减100、银两减10万）</font>"&un&"："&msg&"" & "</marquee>"
else
response.write "<script language='javascript'>alert('你连走路都没劲还想造谣惑众?');</script>"
response.end
end if
rst.Close
set rs=nothing
end function
%>