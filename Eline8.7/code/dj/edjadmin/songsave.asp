<!--#include file="conn.asp"-->
<!--#include file="../const.asp"-->
<!--#include file="checkadmin.inc"-->
<%
page=trim(request.querystring("page"))
AskSClassid=trim(request.querystring("AskSClassid"))
LF_Path=request.form("LF_path")
MusicName=request.form("MusicName")
MusicType=request.form("MusicType")
DJUser=request.form("DJUser")
SClassid=request.form("SClassid")
act=request("act")

if act<>"del" and act<>"istop" then
	if SClassid="" then
		response.write("操作失误，请选择<font color=red size=8>分类</font>!<a href=### onclick=window.history.go(-1);>点这里返回</a>")
		response.end
	end if
		if MusicName="" then
		response.write("操作失误，请输入<font color=red size=8>舞曲名称</font>!<a href=### onclick=window.history.go(-1);>点这里返回</a>")
		response.end
	end if
		if LF_Path="" then
		response.write("操作失误，请输入<font color=red size=8>舞曲链接地址</font>!<a href=### onclick=window.history.go(-1);>点这里返回</a>")
		response.end
	end if
end if

set rs=server.createobject("adodb.recordset")
	select case act
	case "add"
		call add()
	case "edit"
		call edit()
	case "del"
		call del()
	case "istop"
		call istop()
	case else
	rs.close
	conn.close
	set conn=nothing
	response.write "操作错误，原因未知!"
	Response.End
end select


sub add()
'执行添加舞曲开始
		sql="select * from MusicDJ where (id is null)" 
		rs.open sql,conn,1,3
		rs.addnew
		rs("LF_Path")=LF_Path
		rs("MusicName")=MusicName
		rs("MusicType")=MusicType
		rs("DJUser")=DJUser
		rs("Sclassid")=Sclassid
		rs("dateandtime")=now()
		rs.update
		rs.close
		call Success
		Response.End 
'添加结束
end sub


sub edit()
'编辑舞曲开始
		sql="select * from MusicDJ where id="&request("id")
		rs.open sql,conn,1,3
		rs("LF_Path")=LF_Path
		rs("MusicName")=MusicName
		rs("MusicType")=MusicType
		rs("DJUser")=DJUser
		rs("Sclassid")=Sclassid
		rs("dateandtime")=now()
		rs.update
		rs.close
		call Success
		Response.End 
		
'编辑结束
end sub


sub istop()
'舞曲置顶
		sql="select istop from MusicDJ where id="&request("id") 
		rs.open sql,conn,1,3
		if not rs.EOF then
			if rs("istop")=0 then
				rs("istop")=1
			else
				rs("istop")=0
			end if
			rs.update
		end if
		rs.close
		Response.Write "舞曲<font color=red size=8>置顶(撤消)</font>操作成功，<a href=### onclick=window.history.go(-1);>点这里返回!</a>"
		Response.End 
'置顶结束
end sub


sub del()
'删除舞曲开始
	sql="delete from MusicDJ where id="&request.QueryString("ID")
	rs.open sql,conn,1,1
	conn.close
	set conn=nothing
	Sclassid=request("Sclassid")
	page=request("page")
	response.redirect "songmana.asp?Sclassid="&Sclassid&"&page="&page&""

'删除结束
end sub




sub Success
%>
<body>

      <div align="center">
        <center>
          <table width="400" border="1" cellspacing="3" cellpadding="3" class="Tableline" style="border-collapse: collapse" bordercolor="#111111">
            <tr>
              <td width="100%" align="center" bgcolor="#FF9933">舞曲<%if act="add" then%>添加<%else%>修改<%end if%>成功</td>
            </tr>
            <tr>
              <td width="100%">舞 曲 名：<%=MusicName%></td>
            </tr>
            <tr>
              <td>制 作 人：DJ 舞吧</td>
            </tr>
            
            <tr>
              <td>播放方式：<%=MusicType%></td>
            </tr>
            
            <tr>
              <td>试听地址：<a href=<%=djserver%><%=LF_Path%> title=试一下链接是否正确><%=djserver%><%=LF_Path%></a></td>
            </tr>
            
            <tr>
              <td height="15" align=center bgcolor="#FF9933">
              <%
              	Sclassid=request("Sclassid")
				page=request("page")
              %>
              <input type="button" name="button1" value="返  回" onclick=window.open("songmana.asp?Sclassid=<%=Sclassid%><%if act<>"add" then%>&page=<%=page%><%end if%>","_self")>&nbsp;<input type="button" name="button2" value="继续<%if act="add" then%>添加<%else%>修改<%end if%>" onclick="javascript:history.go(-1)"></td>
            </tr>
          </table>
        </center>
    </div>

</body>
</html>
<%end sub%>