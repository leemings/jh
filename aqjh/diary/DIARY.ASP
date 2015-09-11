<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
action=request("action")
if action="reg" then
rjname=request("rjname")
sm=request("sm")
password=request("password")
lb1=request("lb1")
lb2=request("lb2")
lb3=request("lb3")
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
 rs1.Open "Select * From 日记用户 where 姓名='"&aqjh_name&"'", conn1, 1,3
 if rs1.eof or rs1.bof then
 rs1.addnew
        end if
 rs1("日记本名")=rjname
 rs1("姓名")=aqjh_name
 rs1("密码")=password
 rs1("说明")=sm
 rs1("建立日期")=now()
 rs1("lb1")=lb1
 rs1("lb2")=lb2
 rs1("lb3")=lb3
 rs1.update
 rs1.close
 set rs1=nothing
 conn1.close
 set conn1=nothing
 Response.Redirect "mydiary.asp"
end if
if action="add" then
   diary1=Request.Form("diary")
   title1=Request.Form("title")
   weather=Request.Form("weather")
   lb=Request.Form("lb")
   mm=Request.Form("mm")
   bmcs=Request.Form("bmcs")
   face=Request.Form("face")
   dd=Request.Form("my_year")&"-"&Request.Form("my_month")&"-"&Request.Form("my_day")
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   rs1.Open "SELECT * FROM 日记",conn1,1,3
   rs1.AddNew
   rs1("用户名")=aqjh_name
   rs1("日记标题")=title1
   rs1("日记内容")=diary1
   rs1("天气")=weather
   rs1("心情")=face
   rs1("lb")=lb
   rs1("保密")=mm
   rs1("保密条件")=bmcs
   rs1("书写日期")=dd
   rs1.Update
   set rs1=nothing
   conn1.close
   set conn1=nothing
   Response.Redirect "mydiary.asp"
end if
if action="del" then
   id=Request.QueryString("id")
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   rs1.Open "select * from 日记 where id="&id,conn1,1,1
   user=rs1("用户名")
   set rs1=nothing
   conn1.close
   set conn1=nothing
   if user<>aqjh_name and aqjh_name<>Application("aqjh_user") then
   Response.Write "<script Language=Javascript>alert('提示：无权删除！！');location.href = 'javascript:history.go(-1)';</script>"
   else
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   rs1.Open "delete * from 日记 where id="&id,conn1,1,3
   Response.Write "<script Language=Javascript>alert('提示：删除成功！！');location.href = 'javascript:history.go(-1)';</script>"
   set rs1=nothing
   conn1.close
   set conn1=nothing
   end if
end if
if action="edit" then
   id=Request("id")
   diary1=Request.Form("diary")
   title1=Request.Form("title")
   weather=Request.Form("weather")
   lb=Request.Form("lb")
   face=Request.Form("face")
   mm=Request.Form("mm")
   bmcs=Request.Form("bmcs")
   dd=Request.Form("my_year")&"-"&Request.Form("my_month")&"-"&Request.Form("my_day")
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   sql1="update 日记 set 日记标题='"&title1&"',日记内容='"&diary1&"',天气="&weather&",心情="&face&",lb="&lb&",保密="&mm&",保密条件='"&bmcs&"',书写日期="&dd&" where id="&id
   rs1.Open sql1,conn1,1,3
   set rs1=nothing
   conn1.close
   set conn1=nothing
   Response.Redirect "mydiary.asp"
end if
if action="ly" then
  id=Request("id")
  titly=Request("titly")
  myly=Request("myly")
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   rs1.Open "SELECT * FROM 留言",conn1,1,3
   rs1.AddNew
   rs1("姓名")=aqjh_name
   rs1("标题")=titly
   rs1("内容")=myly
   rs1("留言id")=id
   rs1("时间")=Date()
   rs1.Update
   set rs1=nothing
   conn1.close
   set conn1=nothing
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   sql1="update 日记 set 留言数=留言数+1"&" where id="&id
   rs1.Open sql1,conn1,1,3
   set rs1=nothing
   conn1.close
   set conn1=nothing
   Response.Write "<script Language=Javascript>alert('提示：留言成功！！');location.href = 'javascript:history.go(-1)';</script>"
end if
if action="lydel" then
   id=Request("id")
   lyid=Request("lyid")
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   rs1.Open "select * from 日记 where id="&id,conn1,1,1
   if rs1("用户名")<>aqjh_name and aqjh_name<>Application("aqjh_user") then
   Response.Write "<script Language=Javascript>alert('提示：无权删除！！');location.href = 'javascript:history.go(-1)';</script>"
   else
       set rs1=nothing
       conn1.close
       set conn1=nothing
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
       sql1="update 日记 set 留言数=留言数-1"&" where id="&id
       rs1.Open sql1,conn1,1,3
       set rs1=nothing
       conn1.close
       set conn1=nothing
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
       rs1.Open "delete * from 留言 where id="&lyid,conn1,1,3
       Response.Write "<script Language=Javascript>alert('提示：删除成功！！');location.href = 'javascript:history.go(-1)';</script>"        
   end if
conn1.close
set conn1=nothing
end if
if action="delly" then
   id=Request("id")
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
   rs1.Open "select * from 日记 where id="&id,conn1,1,1
   if rs1("用户名")<>aqjh_name and aqjh_name<>Application("aqjh_user") then
   Response.Write "<script Language=Javascript>alert('提示：无权删除！！');location.href = 'javascript:history.go(-1)';</script>"
   else
       set rs1=nothing
       conn1.close
       set conn1=nothing
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
       sql1="update 日记 set 留言数=0"&" where id="&id
       rs1.Open sql1,conn1,1,3
       set rs1=nothing
       conn1.close
       set conn1=nothing
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
       rs1.Open "delete * from 留言 where 留言id="&id,conn1,1,3
       Response.Write "<script Language=Javascript>alert('提示：删除成功！！');location.href = 'javascript:history.go(-1)';</script>"  
    end if
conn1.close
set conn1=nothing      
end if
%>