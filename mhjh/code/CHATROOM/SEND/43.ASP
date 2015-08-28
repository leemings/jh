<%
function gbzdhf(un,st)
         rst.open "select 自动回复 from 用户 where 姓名='"&un&"'",conn
if rst("自动回复")=true then 
        conn.Execute "update 用户 set 自动回复=false where 姓名='"&un&"' and 会员=true"
	gbzdhf="【关闭自动回复】会员##关闭了自动回复功能"
        rst.close
        set rst=nothing
else
        gbzdhf="【关闭自动回复】##，您尚未启用自动回复无需关闭，如果您还不是会员，请尽快加入会员的行列享受更多功能和服务！"
        rst.close
        set rst=nothing
end if
end function
%>
