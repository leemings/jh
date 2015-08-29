<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject ("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"'",conn
msg="<head><link rel='stylesheet' href='../style.css'><title>自立门派</title></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif' topmargic=100><p align=center><font color=ff0000 size=4 face=隶书>自立门派</font><p align=center><img border=0 src=../image/dragon11.gif  width=160 height=100></P><br>谢谢"&username&"大侠你的光临,这里是快乐江湖的门派创建的地方,如果你已经在魔幻玩了很久,有足够的实力，那么你就可以创立自己的门派了,看看你是否达到条件吧,呵呵。<br>条件:攻击,防御为50万,体力,内力各500万,银两500万,精力10万，等级20级<br>注意：建立成功门派以后是不可以修改的,请慎重填写你的资料<br>您的状态,体力:"&rst("体力")&";内力:"&rst("内力")&";攻击:"&rst("攻击")&";防御:"&rst("防御")&";银两:"&rst("银两")&";精力:"&rst("精力")&";等级:"&rst("等级")
if rst("体力")>5000000 and rst("内力")>1000000 and rst("攻击")>500000 and rst("防御")>500000 and rst("银两")>5000000  and rst("身份")<>"掌门" and rst("门派")<>"官府" and rst("等级")>=20  and rst("精力")>100000 then msg=msg&"<form action=addcorp2.asp method=post><table border=3 width='80%' align=center><tr><td>门派名称</td><td><input type=text name='corpname' value='' maxlength=7 size=7></td></tr><tr><td>门派说明</td><td><input type=text name='corpcond' value=''　size=50 maxlength=100></td><tr><td>启动资金</td><td><input type=text value='1000000' name='corpsilver' size=10 maxlength=10></td></tr></tr><tr><td>加入条件</td><td><input type=radio name='tj1' value='True' checked>所有<br><input type=radio name=tj1 value='False'>全不<br><input type=radio name=tj1 value='Other'>其他<ul><dd>性别<input type=checkbox name=sexchk><input type=radio name=sex value='男' checked>男<input type=radio name=sex value='女'>女<dd>体力<input type=checkbox name='hpchk'><input type=radio name=hpopt value='>='  checked>大于等于<input type=radio name=hpopt value='<='>小于等于<input type=text name=hp value=500 size=8 maxlength=8><dd>攻击<input type=checkbox name='attackchk'><input type=radio name=attackopt value='>='  checked>大于等于<input type=radio name=attackopt value='<='>小于等于<input type=text name=attack value=500 size=8 maxlength=8><dd>防御<input type=checkbox name='defencechk'><input type=radio name=defenceopt value='>='  checked>大于等于<input type=radio name=defenceopt value='<='>小于等于<input type=text name=defence value=500 size=8 maxlength=8><dd>道德<input type=checkbox name='moralchk'><input type=radio name=moralopt value='>='  checked>大于等于<input type=radio name=moralopt value='<='>小于等于<input type=text name=moral value=500 size=8 maxlength=8><dd>银两<input type=checkbox name='silverchk'><input type=radio name=silveropt value='>='  checked>大于等于<input type=radio name=silveropt value='<='>小于等于<input type=text name=silver value=500 size=8 maxlength=8></ul></td></tr><tr><td colspan=2 align=center><input type=submit value='自立门派'></td></tr></table></form>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg&"</body>"
%>
