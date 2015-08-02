<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.redirect "../error.asp?id=016"
corpcond=server.HTMLEncode(Trim(Request.Form("corpcond")))&",加入条件："
if instr(corpcond,"'")<>0 or instr(corpcond,chr(34))<>0 then Response.Redirect "../error.asp?id=056"
tj1=Request.Form("tj1")
if tj1="True" then
	tj="and True"
	corpcond=corpcond&"全部"
elseif tj1="False" then
	tj="and False"
	corpcond=corpcond&"拒绝所有"
else
	if Request.Form("sexchk")="on"  then 
		tj=tj&"and 性别='"&Request.Form("sex")&"' "
		corpcond=corpcond&"性别："&Request.Form("sex")&";"
	end if	
	if Request.Form("moralchk")="on" then 
		tj=tj&"and 道德"&Request.Form("moralopt")&Request.Form("moral")&" "
		corpcond=corpcond&"道德："&Request.Form("moralopt")&Request.Form("moral")&";"
	end if	
	if Request.Form("attackchk")="on" then 
		tj=tj&"and 攻击"&Request.Form("attackopt")&Request.Form("attack")&" "
		corpcond=corpcond&"攻击"&Request.Form("attackopt")&Request.Form("attack")&";"
	end if
	if Request.Form("defencechk")="on" then 
		tj=tj&"and 防御"&Request.Form("defenceopt")&Request.Form("defence")&" "
		corpcond=corpcond&"防御"&Request.Form("defenceopt")&Request.Form("defence")&";"
	end if
	if Request.Form("hpchk")="on" then 
		tj=tj&"and 体力"&Request.Form("hpopt")&Request.Form("hp")&" "
		corpcond=corpcond&"体力"&Request.Form("hpopt")&Request.Form("hp")&";"
	end if
	if Request.Form("silverchk")="on" then 
		tj=tj&"and 银两"&Request.Form("silveropt")&Request.Form("silver")&" "
		corpcond=corpcond&"体力"&Request.Form("hpopt")&Request.Form("hp")&";"
	end if
end if
tj=mid(tj,5)
corpname=server.HTMLEncode(Trim(Request.Form("corpname")))
corpsilver=server.HTMLEncode(Trim(Request.Form("corpsilver")))
if not isnumeric(corpsilver) then Response.Redirect "../error.asp?id=024"
corpsilver=clng(corpsilver)
silver=corpsilver\100000
if silver<0 then 
	silver=0
elseif silver>1000 then
	silver=1000
end if
if corpname="" or corpname="无" or instr(corpname,"'")<>0 or instr(corpname,chr(34))<>0 then Response.Redirect "../error.asp?id=056"
if tj="" then Response.Redirect "../error.asp?id=102"
%>
<head><link rel='stylesheet' href='../chatroom/style1.css'></head><body oncontextmenu=self.event.returnValue=false background='<%=bgimage%>' bgcolor='<%=bgcolor%>' topmargin=100><p align=center><font color=ff0000 size=5>自立门户</font></p><hr>
<!--恭喜恭喜，经过长期的奋斗,你终于建立了自己的门派<font color=0000ff size=6><%=corpname%></font> 精力减少10000,银两减少<%=corpsilver%><br><p align=center><a href=javascript:location.replace('welcome.asp');>返回</a></p>-->
<form action=addcorp3.asp method=post>
<table border=3 align=center width=80%>
<tr><td>门&nbsp;&nbsp;&nbsp;&nbsp;派</td><td><input type=text name=corpname value="<%=corpname%>" size=50 readonly></td></tr>
<tr><td>简&nbsp;&nbsp;&nbsp;&nbsp;介</td><td><input type=text name=corpcond value="<%=corpcond%>" size=50 readonly></td></tr>
<tr><td>工资系数</td><td><input type=text name=corpsilver value="<%=silver%>" size=50 readonly></td></tr>
<tr><td>加入条件</td><td><input type=text name=corptj value="<%=tj%>" size=50 readonly></td></tr>
<tr><td colspan=2 align=center><input type=submit value="开始建立"> <input type=button value="重新设定" onclick="javascript:history.back();"></td></tr>
</table>
</form>