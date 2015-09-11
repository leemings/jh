<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
Response.Write "<body>"
dim deld
deld=trim(Request.form ("deld"))
if deld<>"" then 
	call delok()
else
	call deljd()
end if
sub delok()
%>
<!--#include file="dadata.asp"-->
<%
sql="SELECT id from 宴会列表 where DateDiff('d',时间,date())>" & deld
Set Rs=connt.Execute(sql)

do while not rs.bof and not rs.eof
id=rs(0)
	sql1 = "delete * from 宴会列表 where ID=" & id
	connt.execute sql1
	sql1 = "delete * from 宴会者 where 宴会名=" & id
	connt.execute sql1
	rs.movenext
loop
Rs.close
connt.Close
Set Rs=Nothing
Set connt = Nothing
%>
<LINK href=../aqjh_admin/css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table width="75%" border="1" align="center" cellpadding="5" cellspacing="5">
  <tr>
    <td align="center" >所有开宴时间超过的酒宴都被清除了，建议在数据设置中</td>
  </tr>
</table>
<%
callexit()
End Sub
sub deljd()
%><LINK href=../aqjh_admin/css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p>使用这个功能，你可以清理酒店里过期的酒席，节约数据库空间提高性能<br>
  </font>方便玩家吃洒宴（大多不能吃的洒宴会令人眼花）<br>
  </font>清除成功会同时在聊天室发布公告</font></p>
<table width="75%" border="1" align="center" cellpadding="5" cellspacing="5">
  <tr> 
    <td align="center" ><form name="form1" method="post" action="CL.ASP">
清除开宴时间超过： 
        <select name="deld">
          <option value="30" selected>30</option>
          <option value="20">20</option>
          <option value="10">10</option>
          <option value="5">5</option>
        </select>
        天的酒宴！</font> 
        <input type="submit" name="Submit" value="·清 理·">
      </form></td>
  </tr>
</table>
<%
end sub
sub callexit()
says="<font color=black>【清理酒店】</font><font color=red>站长[ " & aqjh_name &" ]对爱情江湖酒店进行了清理，所有超过[ " & deld &" ]天的洒宴都被删除了！</font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
call addmsg(saystr)
end sub
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>