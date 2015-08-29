<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=trim(Request.form("name"))
mess=trim(Request.form("mess"))
mess=LCase(trim(Request.form("mess")))
mess=replace(mess,"'","")
mess=replace(mess,chr(34),"")
mess=Replace(mess,"<","")
mess=Replace(mess,">","")
mess=Replace(mess,"\x3c","")
mess=Replace(mess,"\x3e","")
mess=Replace(mess,"\074","")
mess=Replace(mess,"\74","")
mess=Replace(mess,"\75","")
mess=Replace(mess,"\76","")
mess=Replace(mess,"&lt","")
mess=Replace(mess,"&gt","")
mess=Replace(mess,"\076","")
badstr="江湖总站|世纪江湖|天外天江湖|六六江湖|国忠|清评|文钢|炳川|朝阳|tama|ta ma|tamu|ta mu|学子天地|破江湖|玩别的江湖|去别的江湖|去我的江湖|来我的江湖|玩我的江湖|去其它江湖|去其他江湖|玩其他江湖|玩其它江湖|烂江湖|滥江湖|射精|奸|我干|官府|鸟人|发全|喀什|咔什|cfq|爬你达来蛋|撅你达来蛋|死你达来蛋|达来蛋|chenfaquan|chen fa quan|chen|ma de|kan ni|fa|quan|ni ma|mao|ni mu|niba|ni ba|nima|nimu|ni nai nai|ni jie|ni mei|nai|ye|ma|ba|zu|zhu|傻鸟|站掌|占长|战长|丈长|帐长|我塞|我靠|老师|砍你|砍了你|砍死你|1线|捅你|捅死你|教师|妓女|陈发|做鸭|作鸭|紫华|sai|管理员|网管|kao|cao|yixiantian|yi xian tian|yxt|eline|塞你母|8196316|13808556346|一线|一线天|线天|一钱天|你母|你木|发铨|去死|站长|老大|吃屎|妈|爸|爹|娘|日你|尻|爷爷|奶奶|你爷|你奶|他爷|他奶|操你|我操|干死你|王八|逼|傻|贱|剁了你|做婊|做表|作婊|作表|婊|表子|靠你|叉|插你|插死|干你|干死|日死|鸡巴|睾丸|包皮|龟头|屄|赑|妣|肏|奶子|尻|屌|作爱|做爱|床上|抱抱|鸡八|处女|处男|打炮|爸|我儿|麻痹|吗比|叼|草你|阿冒|阿茂|啊冒|啊茂|老吗|乌龟|屎|屁|笨蛋|白痴|滚|江湖人|大便|死八公|http|www|com|cn|org|net|//|fuck|cao|gan|sai|你吗"
bad=split(badstr,"|")
for i=0 to UBound(bad)
mess=Replace(mess,bad(i),"**")
next
money=int(abs(Request.form("money")))
if name=sjjh_name or name="" or mess="" then 
	Response.Write "<script Language=Javascript>alert('提示：姓名不能为空，不能杀自己？！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if money<10000 or money>100000000 then
	Response.Write "<script Language=Javascript>alert('提示：你有这么多的钱嘛？！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM h WHERE a='"&sjjh_name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你要杀的人不存在。。！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
'校验用户 魅力大于100，钱大于1000
rs.open "SELECT * FROM 用户 WHERE 姓名='"&sjjh_name&"'"&" and 银两>"&money,conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你有这么多的钱嘛？！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT * FROM e WHERE a='"&sjjh_name&"' and b='"&name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你已经申请过了不能重复操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
	conn.execute "update 用户 set 银两=银两-"&money&" where 姓名='"&sjjh_name&"'"
	conn.Execute "INSERT INTO e (a,b,c,d,e,f) VALUES('"&sjjh_name&"','"&name&"','"&mess&"',now(),"&money&",'无')"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：数据提交成功！');location.href = 'killer.asp';</script>"
response.end
%>