<!--#include file="conn.asp"-->
<!--#include file="checkadmin.inc"-->
<%
SClassid=request.querystring("SClassid")
Classid=request.form("Classid")
Sclass=request.form("Sclass")
act=request("act")

if act<>"del" and act<>"add" then
	if SClassid="" then
		response.write("操作失误，输入<font color=red size=8>二级分类名</font>!<a href=### onclick=window.history.go(-1);>点这里返回</a>")
		response.end
	end if

		if ClassID="" then
		response.write("操作失误，请选择此二级分类所属的<font color=red size=8>一级分类</font>!<a href=### onclick=window.history.go(-1);>点这里返回</a>")
		response.end
	end if
end if
if act="add" or act="edit" then
set rs=server.createobject("adodb.recordset")
sql="select * from Class where Classid="&Classid
rs.open sql,conn,1,1
if not rs.eof then
ClassName=rs("class")
else
		response.write("操作失误，没有ID为  "&Classid&"  的<font color=red size=8>一级分类</font>!<a href=### onclick=window.history.go(-1);>点这里返回</a>")
		response.end
end if
rs.close
set rs=nothing
end if

set rs=server.createobject("adodb.recordset")
	select case act
	case "add"
		call add()
	case "edit"
		call edit()
	case "del"
		call del()
	case else
	rs.close
	conn.close
	set conn=nothing
	response.write "操作错误，原因未知!"
	Response.End
end select


sub add()
'执行添加二级分类开始
		sql="select * from SClass where (SClassid is null)" 
		rs.open sql,conn,1,3
		rs.addnew
		rs("Sclass")=Sclass
		rs("Classid")=Classid
		rs.update
		rs.close
		call Success
		Response.End 
'添加结束
end sub


sub edit()
'编辑二级分类开始
		sql="select * from SClass where SClassid="&SClassid
		rs.open sql,conn,1,3
		rs("Sclass")=Sclass
		rs("Classid")=Classid
		rs.update
		rs.close
		call Success
		Response.End 
		
'编辑结束
end sub


sub del()
'删除二级分类开始
	sql="delete from SClass where SClassid="&request.QueryString("SClassID")
	rs.open sql,conn,1,1
	conn.close
	set conn=nothing
	response.redirect "SClassmana.asp"

'删除结束
end sub




sub Success
%>
<body>

      <div align="center">
        <center>
          <table width="400" border="1" cellspacing="3" cellpadding="3" class="Tableline" style="border-collapse: collapse" bordercolor="#111111">
            <tr>
              <td width="100%" align="center" bgcolor="#FF9933">二级分类<%if act="add" then%>添加<%else%>修改<%end if%>成功</td>
            </tr>
            <tr>
              <td width="100%">二级分类名：<%=SClass%></td>
            </tr>
            
            <tr>
              <td>所属一级分类：<%=ClassName%> </td>
            </tr>
            
            <tr>
              <td height="15" align=center bgcolor="#FF9933">
              <%
              	Sclassid=request("Sclassid")
				page=request("page")
              %>
              <input type="button" name="button1" value="返  回" onclick=window.open("SClassmana.asp","_self")>&nbsp;<input type="button" name="button2" value="继续<%if act="add" then%>添加<%else%>修改<%end if%>" onclick="javascript:history.go(-1)"></td>
            </tr>
          </table>
        </center>
    </div>

</body>
</html>
<%end sub%>