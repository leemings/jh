<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'寻找暗器♀爱情江湖→wWw.7758530.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if

act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找暗器】<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找暗器
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 等级,操作时间,会员等级,金币 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
dj=rs("等级")

fla=rs("金币")
if rs("金币")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的金币不够无法施展呀，至少也得2个啊！');window.close();}</script>"
	response.end
end if
if dj<60 then
Response.Write "<script language=JavaScript>{alert('此功能需要60级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
hydj=rs("会员等级")
if hydj<3 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：会员等级为3级才可以进行寻找暗器！购买2级会员赠送200级每天可以领到100个金币！购买加QQ51726805 或打电话：13635716119');}</script>"
	Response.End
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=30-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
f=Minute(time())
if f<40 or f>50 then
	Response.Write "<script language=JavaScript>{alert('暗器铺现在没有开放，寻找暗器时间为每小时的40-50分钟！');window.close();}</script>"
	Response.End 
end if
if rs("职业")<>"冒险家" then
	Response.Write "<script language=JavaScript>{alert('你非冒险家，不能找宝物！！请去职业转换为冒险家！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
conn.execute "update 用户 set 金币=金币-2,操作时间=now()  where 姓名='"&aqjh_name&"'"
randomize 
r=int(rnd*10)+1
select case r
case 1
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>铁蒺藜</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"铁蒺藜",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close


case 2
xunfaqi=aqjh_name & "你在" & to1 & "的家里做客，碰到了永不放弃也在那里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>柳叶镖</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"柳叶镖",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close

	
case 3
xunfaqi=aqjh_name & "到永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>梅花镖</font></b>]一个>。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"梅花镖",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "到永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>袖箭</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"袖箭",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "到永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>毒镖</font></b>]一个"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"毒镖",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 6
xunfaqi=aqjh_name & "到永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>铜钱镖</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"铜钱镖",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 7
xunfaqi=aqjh_name & "到永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>毒龙镖</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"毒龙镖",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 8
xunfaqi=aqjh_name & "到永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>袖剑</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"袖剑",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "到永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>毒针</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"毒针",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "永不放弃家里聚会。收到站长<font color=#0000FF>永不放弃</font>赠送给您礼物[<b><font color=red>飞蝗石</font></b>]一个。"
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	anqi=add(rs("w4"),"飞蝗石",1)
	conn.execute "update 用户 set  w4='"&anqi&"' where 姓名='"&aqjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
