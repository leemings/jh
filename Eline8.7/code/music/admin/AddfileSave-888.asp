<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<%
act=request("act")
Page=request.QueryString("Page")
Askclassid=request.QueryString("Askclassid")
Asksclassid=request.QueryString("Asksclassid")
AskNclassid=request.QueryString("AskNclassid")
Specialid=request.QueryString("Specialid")
if act<>"SetIsGood" then
	Classic=request.form("Classic")
	ConClassic=split(Classic,",")
	Classid=ConClassic(0)
	SClassid=ConClassic(1)
	NClassid=ConClassic(2)
end if
set rs=server.createobject("adodb.recordset")

founerr=false
if act<>"SetIsGood" then
	if trim(Request.Form ("name"))="" then
		founderr=true
		errmsg="<li>商品名称不能为空</li>"
	else
		Name=trim(Request.Form ("Name"))
	end if

	sql="select Sclass from Sclass where Sclassid="&SClassid
	rs.open sql,conn,1,1
	if not rs.EOF then
		Sclass=rs("Sclass")
	else
		founderr=true
		errmsg=errmsg+"<li>二级分类选择错误</li>"
	end if
	rs.close

	sql="select Nclass from Nclass where Nclassid="&NClassid
	rs.open sql,conn,1,1
	if not rs.EOF then
		Nclass=rs("Nclass")
	else
		founderr=true
		errmsg=errmsg+"<li>三级分类选择错误</li>"
	end if
	rs.close
end if

if founderr=true then
	call error()
else
	if act<>"SetIsGood" then
		Name=request("name")
		if trim(Request.Form("Yuyan"))="" then	
			Yuyan="未知"
		else
			Yuyan=request.form("Yuyan")	
		end if
		if trim(Request.Form("Gongsi"))="" then
			Gongsi="未知"
		else
			Gongsi=request.form("Gongsi")
		end if
		if trim(Request.Form("pic"))="" then
			pic="images/Nophoto.gif"
		else
			pic=request.form("pic")		
		end if
		if trim(Request.Form("intro"))="" then
			intro="暂无"
		else
			intro=htmlencode2(request.form("intro"))
		end if
	end if

	if act="edit" and Specialid<>"" then
'修改商品资料
		sql="select * from Special where Specialid="&Specialid 
		rs.open sql,conn,1,3
		rs("name")=trim(name)
		rs("Yuyan")=trim(Yuyan)
		rs("Gongsi")=trim(Gongsi)
		rs("pic")=trim(pic)
		rs("Classid")=Classid
		rs("SClassid")=SClassid
		rs("SClass")=SClass
		rs("Nclassid")=Nclassid
		rs("Nclass")=Nclass
		rs("intro")=htmlencode2(trim(intro))
		rs.update
		rs.close
'结束修改

	elseif act="add" then
		sql="select * from Special where (Specialid is null)" 
		rs.open sql,conn,1,3
		rs.addnew
		rs("name")=trim(name)
		rs("Yuyan")=trim(Yuyan)
		rs("Gongsi")=trim(Gongsi)
		rs("pic")=trim(pic)
		rs("classid")=classid
		rs("SClass")=SClass
		rs("Sclassid")=Sclassid
		rs("Nclassid")=Nclassid
		rs("Nclass")=Nclass
		rs("intro")=htmlencode2(trim(intro))
		rs("Time")=date
		rs.update
		rs.close

	elseif act="SetIsGood" then
'		sql="select IsGood from Special where IsGood=true"
'		rs.open sql,conn,1,3
'		if not rs.EOF then
'			do while not rs.EOF
'				rs("IsGood")=false
'				rs.update
'			rs.MoveNext 
'			loop
'		end if
'		rs.close
		sql="select IsGood from Special where Specialid="&Specialid
		rs.open sql,conn,1,3
		if not rs.EOF then
			if rs("IsGood")=true then
				rs("IsGood")=false
			else
				rs("IsGood")=true
			end if
			rs.update
		end if
		rs.close
	else
		conn.close
		set conn=nothing
		Response.Write "操作错误！"
		Response.End 
	end if
	set rs=nothing

end if
conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="RSHOP" content="http://www.119music.com">
<meta Author="Recall Star" content="音乐听吧">
<title>音乐听吧</title>
<!--#include file="style.asp"-->
</head>
<body topmargin="111" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellspacing="0" width="60%">
    <tr>
      <td width="100%" bgcolor="#CC0066">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td width="100%" bgcolor="#FFFFFF" height="80" align="center">
                <b>O K !&nbsp; 操 作 完 成 !&nbsp; ^_^</b>
                <p><b><a href="javascript:history.go(-1)">...::: 点 此 返 回 
                :::...</a></b>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
</body>                    
</html>           



