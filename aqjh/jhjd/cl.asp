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
sql="SELECT id from ����б� where DateDiff('d',ʱ��,date())>" & deld
Set Rs=connt.Execute(sql)

do while not rs.bof and not rs.eof
id=rs(0)
	sql1 = "delete * from ����б� where ID=" & id
	connt.execute sql1
	sql1 = "delete * from ����� where �����=" & id
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
    <td align="center" >���п���ʱ�䳬���ľ��綼������ˣ�����������������</td>
  </tr>
</table>
<%
callexit()
End Sub
sub deljd()
%><LINK href=../aqjh_admin/css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p>ʹ��������ܣ����������Ƶ�����ڵľ�ϯ����Լ���ݿ�ռ��������<br>
  </font>������ҳ����磨��಻�ܳԵ�����������ۻ���<br>
  </font>����ɹ���ͬʱ�������ҷ�������</font></p>
<table width="75%" border="1" align="center" cellpadding="5" cellspacing="5">
  <tr> 
    <td align="center" ><form name="form1" method="post" action="CL.ASP">
�������ʱ�䳬���� 
        <select name="deld">
          <option value="30" selected>30</option>
          <option value="20">20</option>
          <option value="10">10</option>
          <option value="5">5</option>
        </select>
        ��ľ��磡</font> 
        <input type="submit" name="Submit" value="���� ��">
      </form></td>
  </tr>
</table>
<%
end sub
sub callexit()
says="<font color=black>������Ƶ꡿</font><font color=red>վ��[ " & aqjh_name &" ]�԰��齭���Ƶ�������������г���[ " & deld &" ]������綼��ɾ���ˣ�</font>"			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
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