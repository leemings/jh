<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_grade<>10 then Response.Redirect "../error.asp?id=439"
if session("aqjh_adminok")<>true then response.end
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
act=sqlstr(request("act"))
Select Case act
case "addnpc"
	call zzz()
	npcname=sqlstr(request.form("npcname"))
	npcsex=sqlstr(request.form("npcsex"))
	npcvalue=sqlstr(request.form("npcvalue"))
	npcimg=sqlstr(request.form("npcimg"))
	npcmoney=sqlstr(request.form("npcmoney"))
	npcgj=sqlstr(request.form("npcgj"))
	npcfy=sqlstr(request.form("npcfy"))
	npcwg=sqlstr(request.form("npcwg"))
	npctl=sqlstr(request.form("npctl"))
	npcnl=sqlstr(request.form("npcnl"))
	npcdie=sqlstr(request.form("npcdie"))
	npcgjl=sqlstr(request.form("npcgjl"))
	npccxl=sqlstr(request.form("npccxl"))
	npcwp=sqlstr(request.form("npcwp"))
	npcwplx=sqlstr(request.form("npcwplx"))
	npcccc=sqlstr(request.form("npcccc"))
	rs.open "Select * from npc Where n姓名='" & npcname & "'",conn,1,1
	if not rs.eof then
		rs.close
		call smsg("[" & npcname & "]此npc已经存在！","-1")
	end if
	rs.close
	rs.Open "SELECT * FROM npc ",conn,1,3
	rs.AddNew
        rs("n姓名")=npcname
	rs("n性别")=npcsex
	rs("n经验")=npcvalue
	rs("n图片")=npcimg
	rs("n银两")=npcmoney
	rs("n攻击")=npcgj
	rs("n防御")=npcfy
	rs("n武功")=npcwg
	rs("n体力")=npctl
	rs("n内力")=npcnl
	rs("n死次数")=npcdie
	rs("n攻击率")=npcgjl
	rs("n出现率")=npccxl
	rs("n物品")=npcwp
	rs("n物品类型")=npcwplx
	rs("n出场词")=npcccc
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','管理记录','新增NPC:" & npcname & "')"
	rs.close
	call smsg("恭喜,npc添加成功!","npclist.asp")

case "delnpc"
	call zzz()
	npcname=sqlstr(request("npcname"))
	conn.Execute "delete * from npc where n姓名='" & npcname &"'"
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','管理记录','删除npc {id}:[" & id & "]')"
	call smsg("删除成功!","npclist.asp")
case "delczc"
	call zzz()
	sn=sqlstr(request("sn"))
	conn.Execute "delete * from czc where sn='" & sn &"'"
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','管理记录','删除充值卡SN:[" & sn & "]')"
	call smsg("删除成功!","czclist.asp")
case "delallczc"
	call zzz()
	conn.Execute "delete * from czc where cz_use=1"
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','管理记录','删除所以用过充值卡成功!')"
	call smsg("删除成功!","czclist.asp")
case "modiczc"
	call zzz()
	id=clng(request("id"))
	pass=request.form("pass")
	cz_use=sqlstr(request.form("cz_use"))
	sm=sqlstr(request.form("sm"))
	'更新资料!
	rs.Open "SELECT * FROM czc where id=" & id,conn,1,3
	if len(pass)=6 then rs("pass")=md5(pass)
	if cz_use=1 then
		rs("cz_use")=1
	else
		rs("cz_use")=0
	end if
	rs("remark")=sm
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','管理记录','修改充值卡ID:" & id & "')"
	rs.close
	call smsg("恭喜,此充值卡资料修改完成!","czclist.asp")
case "modinpc"
	call zzz()
	id=clng(request("id"))
	npcname=sqlstr(request.form("npcname"))
	npcsex=sqlstr(request.form("npcsex"))
	npcvalue=sqlstr(request.form("npcvalue"))
	npcimg=sqlstr(request.form("npcimg"))
	npcmoney=sqlstr(request.form("npcmoney"))
	npcgj=sqlstr(request.form("npcgj"))
	npcfy=sqlstr(request.form("npcfy"))
	npcwg=sqlstr(request.form("npcwg"))
	npctl=sqlstr(request.form("npctl"))
	npcnl=sqlstr(request.form("npcnl"))
	npcdie=sqlstr(request.form("npcdie"))
	npcgjl=sqlstr(request.form("npcgjl"))
	npccxl=sqlstr(request.form("npccxl"))
	npczhuren=sqlstr(request.form("npczhuren"))
	npcdiren=sqlstr(request.form("npcdiren"))
	npcwp=sqlstr(request.form("npcwp"))
	npcwplx=sqlstr(request.form("npcwplx"))
	npcccc=sqlstr(request.form("npcccc"))
	'更新资料!
	rs.Open "SELECT * FROM npc where id=" & id,conn,1,3
        if rs.eof or rs.bof then
		rs.close
		call smsg("用户不存在!","-1")
	end if
	rs("n姓名")=npcname
	rs("n性别")=npcsex
	rs("n经验")=npcvalue
	rs("n图片")=npcimg
	rs("n银两")=npcmoney
	rs("n攻击")=npcgj
	rs("n防御")=npcfy
	rs("n武功")=npcwg
	rs("n体力")=npctl
	rs("n内力")=npcnl
	rs("n死次数")=npcdie
	rs("n攻击率")=npcgjl
	rs("n出现率")=npccxl
	rs("n主人")=npczhuren
	rs("n敌人")=npcdiren
	rs("n物品")=npcwp
	rs("n物品类型")=npcwplx
	rs("n出场词")=npcccc
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','管理记录','修改NPC人物：[" & id & "]')"
	rs.close
	call smsg("恭喜,此NPC资料修改完成!","npclist.asp")
End Select
sub zzz
if aqjh_name<>Application("aqjh_user") then 
	call smsg("你不是正站长，你不能操作!","-1")
end if
end sub
function sqlstr(mystr)
mystr=lcase(trim(mystr))
mystr=Replace(mystr,"union","u_nion")
mystr=Replace(mystr,"'","’")
mystr=Replace(mystr,"not","n_o_t")
if mystr="" then
	call smsg("输入的数据为空不能操作!","-1")
end if
sqlstr=mystr
end function
Sub smsg(str1, str2)
 set rs=nothing
 conn.close
 set conn=nothing
 onend = 0
 Select Case str2
 Case "-1","-2","-3" '回退
    str2 = "location.href = 'javascript:history.go(-1)';"
 Case "1" '关闭
    str2 = "window.close();"
 Case "2" '不返回
    str2 = ""
 Case "3" '不结
    str2 = ""
    onend = 1
 Case Else '回退到指定文件
    str2 = "location.href = '" & str2 & "';"
 End Select
 Response.Write "<script language=JavaScript>{alert('提示：" & str1 & "');" & str2 & "}</script>"
 If onend = 0 Then
    Response.End
 End If
End Sub
%>
