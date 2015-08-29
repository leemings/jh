<%
dim Database,CheckLogin,LoginURL,DatabaseMember,DatabaseUserName,DatabaseCash,BetLimit
CheckLogin=request.cookies("aspsky")("username")	'用什么方式记录登录，比如填，不要用""
LoginURL="index.asp"				'登录页面的地址
Database="data/eline_bbs_6.3.0.asp"				'数据库地址
DatabaseMember="user"		'保存用户数据的表名
DatabaseUserName="UserName"		'数据库中用户名的字段名称
DatabaseCash="userWealth"			'数据库中货币的字段名称
BetLimit=100				'每种水果的最高下注金额
%>