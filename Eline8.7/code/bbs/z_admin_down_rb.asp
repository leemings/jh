<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<html>
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%dim str
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	if Request("action") = "backupdata" then
		call backupdata()
	elseif Request("action") = "restoredata" then
		call restoredata()
	elseif Request("action") = "comp" then
		call comp()
	else
		Errmsg=Errmsg+"<br>"+"<li>��������"
		call dvbbs_error()
	end if
end if
conndown.close
set conndown=nothing%>
</Body>

<%Function CheckDir(FolderPath)
	dim fso1
	folderpath=Server.MapPath(".")&"\"&folderpath
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	If fso1.FolderExists(FolderPath) then
		CheckDir = True
	Else
		CheckDir = False
	End if
	Set fso1 = nothing
End Function

Function MakeNewsDir(foldername)
	dim fso1,f
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	Set f = fso1.CreateFolder(foldername)
	MakeNewsDir = True
	Set fso1 = nothing
End Function

sub backupdata()
	dim Dbpath,bkfolder,bkdbname,Fso
	Dbpath=request.form("Dbpath")
	Dbpath=server.mappath(Dbpath)
	bkfolder=request.form("bkfolder")
	bkdbname=request.form("bkdbname")
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(dbpath) then
		If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
		else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
		end if
		response.write "���ݿⱸ����ɣ����������������<br>����ʹ�� FTP ���߽����ݿⱸ�ݣ��Ա�֤���ݰ�ȫ"
	else
		response.write "�Ҳ���������Ҫ���ݵ��ļ���"
	end if
end sub

sub restoredata()
	dim dbpath,backpath,fso
	Dbpath=request.form("Dbpath")
	backpath=request.form("backpath")
	if dbpath="" then
		response.write "��������Ҫ�ָ��ɵ����ݿ���������"
	else
		Dbpath=server.mappath(Dbpath)
	end if
	backpath=server.mappath(backpath)
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(dbpath) then  
		fso.copyfile Dbpath,Backpath
		response.write "���ݿ�ָ���ɣ����������������"
	else
		response.write "����Ŀ¼��û���ҵ����ı����ļ���"
	end if
end sub

sub comp()
	Dim dbpath,boolIs97
	dbpath = request("dbpath")
	boolIs97 = request("boolIs97")
	If dbpath <> "" Then
		dbpath = server.mappath(dbpath)
		response.write(CompactDB(dbpath,boolIs97))
	End If
end sub

Const JET_3X = 4

Function CompactDB(dbPath, boolIs97)
	Dim fso, Engine, strDBPath
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(dbPath) Then
		Set Engine = CreateObject("JRO.JetEngine")
		If boolIs97 = "True" Then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
				"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
				& "Jet OLEDB:Engine Type=" & JET_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
				"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
		End If
		fso.CopyFile strDBPath & "temp.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		Set fso = nothing
		Set Engine = nothing
		CompactDB = "������ݿ�, " & dbpath & ", �Ѿ�ѹ���ɹ�!" & vbCrLf
	Else
		CompactDB = "���ݿ����ƻ�·������ȷ. ������!" & vbCrLf
	End If
End Function%>