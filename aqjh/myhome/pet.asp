<%@ LANGUAGE=VBScript codepage ="936"%>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

pet0=8		'宠物照看的周期,默认为8小时
pet1=1		'每八小时扣的点数(攻,防,心,精气)
pet111=2		'每八小时扣的体力
pet2=10		'喂养时花增加体力
pet3=30		'喂养时种花增加体力数

pet4=5		'锻练时增加的:攻,防
pet41=5		'锻练时增加的:精气
pet5=2		'锻练时咸少的:心情
pet6=15	'锻练时咸少的:体力

pet7=10		'陪伴时增加的:心情
pet8=2		'陪伴时增加的:精气
pet9=15	'陪伴时咸少的:体力

pet10=3		'洗澡时增加的:心情
pet11=2		'洗澡时增加的:攻,防
pet12=15	'洗澡时咸少的:体力
pet121=3		'洗澡时咸少的:精气

pet13=8		'教训时增加的:攻,防
pet14=3	'教训时增加的:精气
pet15=10	'教训时增加的:心情
pet16=15	'教训时咸少的:体力

pet17=20	'休息时增加的:体力
pet18=2		'休息时增加的:精气
pet19=2		'休息时增加的:心情
pet20=2	'休息时增加的:攻,防
pet21=1000000	'打针时所需的银两
pet22=15		'宠物多少天增加一个照料周期15天


set xajh=Server.CreateObject("anjh70.xajh")
call xajh.pet(pet0,pet1,pet111,pet2,pet3,pet4,pet41,pet5,pet6,pet7,pet8,pet9,pet10,pet11,pet12,pet121,pet13,pet14,pet15,pet16,pet17,pet18,pet19,pet20,pet21,pet22)
set xajh=nothing
'call pet()

%>