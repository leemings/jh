<%
function present(un,gr,st,mg)
if gr<10 then
       Response.Write "<script language=JavaScript>{alert('你等级小于10级,新人啊,连自己都照顾不了还想给人东西?省省吧！');}</script>"
       Response.End
       exit function
end if	
rst.Open "select * from 物品  where 所有者='"&un&"' and 数量>0 and 名称='"&mg&"'",conn
if rst.EOF or rst.BOF then
present="<font color=FF0000>【赠送】</font>##形容落魄，身无长物，居然也想拿"&mg&"送给%%。"
else
set rst2=server.CreateObject("adodb.recordset")
rst2.Open "select * from 物品 where 所有者='"&st&"' and 名称='"&mg&"'",conn
if rst2.EOF or rst2.BOF then
conn.execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,数量,价格,所有者,寄售,装备) values('"&mg&"','"&rst("属性")&"',"&rst("体力")&","&rst("内力")&","&rst("攻击")&","&rst("防御")&",'"&rst("特效")&"',1,"&rst("价格")&",'"&st&"',false,false)"
else
conn.execute "update 物品 set 数量=数量+1 where 所有者='"&st&"' and 名称='"&mg&"'"
end if
conn.execute "update 物品 set 数量=数量-1 where 所有者='"&un&"' and 名称='"&mg&"'"
rst2.Close
set rst2=nothing
present="<font color=FF0000>【赠送】</font>##将"&mg&"大大方方的递到%%面前说：“一点心意，菲薄难堪！”"
end if
rst.Close
end function
%>
