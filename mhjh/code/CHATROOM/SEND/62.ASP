<%
'加入天网
function jrtw(un,co,gr)  
if gr<21 then
Response.Write "<script language=JavaScript>{alert('你的等级太低了，要21级才能加入天网组织！');}</script>"
 Response.End
 exit function
end if	
if co<>"无" then
Response.Write "<script language=JavaScript>{alert('你还是一个有门派的人,先离开你的门派,再来加入我们吧！');}</script>"
 Response.End
 exit function
end if
rst.Open "select 攻击,防御,银两 from 用户 where 姓名='"&un&"'",conn
if rst("等级")<21 or rst("攻击")<1000000 or rst("会员")=True or rst("银两")<2000000 then
Response.Write "<script language=JavaScript>{alert('会员才能加入,等级21级,攻击100万,银两200万以上才能申请杀手');}</script>"
Response.End
exit function
end if	
conn.execute "update 用户 set 门派='天网',身份='杀手',银两=银两-2000000 where 姓名='"&un&"'"
jrtw="【天网】##,你交了200万两后成功加入了我们庞大的天网杀人组织，现在你的任务就是收人钱财，替人消灾[img]181.gif[/img]"
session("yx8_mhjh_usercorp")="天网"		
end function
%>

or rst("会员")=True