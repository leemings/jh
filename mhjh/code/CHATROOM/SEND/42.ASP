<%
function kqzdhf(un,st,mg)
         rst.open "select 会员 from 用户 where 姓名='"&un&"'",conn
         if rst("会员")=true then
         kqzdhf="【开启自动回复】会员##启用了自动回复功能"
         conn.Execute "update 用户 set 自动回复=true,回复语句='"&mg&"' where 姓名='"&un&"' and 会员=true"
         rst.close
         set rst=nothing
         else
         kqzdhf="【开启自动回复】##，您启用自动回复功能失败，原因您不是会员，请尽快加入会员的行列享受更多功能和服务！"
         rst.close
         set rst=nothing
         end if
end function
%>
