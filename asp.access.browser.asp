<%
 on error resume next
 title = "asp Db browser For access"
%><!doctype html>
<html lang="zh-CN">
<head>
<meta charset="gb2312" />
<meta name="viewport" content="width=device-width,minimum-scale=1.0,maximum-scale=1.0" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<title><%=title%></title>
<meta name="author" content="yujianyue, admin@ewuyi.net">
<meta name="copyright" content="www.12391.net">
<style>
*{margin:0; padding:0; font-size:12px;}
h1{ font-size:18px;}
a{padding:2px 3px;color:blue;line-height:150%;text-decoration:none;}
a:hover{border:0;background-color:#0180CF;color:white;}
b{padding:5px 10px;background-color:#0180CF;color:white;}
select{width:99%;line-height:150%;}
table{margin:5px auto;border-left:1px solid #a2c6d3;border-top:3px solid #0180CF;width:96vw;}
table td{border-right:1px solid #a2c6d3;border-bottom:1px solid #a2c6d3;}
table td{padding:5px;word-wrap:break-word;word-break:break-all;}
.tt{background:#e5f2fa;line-height:150%;font-size:14px;padding:5px 9px;}
</style>
</head>
<body>
<%

'asp.access.browser V20191215

'功能：列出网站目录下的access数据并浏览：
'1. 遍历文件夹下所有.mdb,.asa文件供选并记忆；
'2. 然后该数据文件下所有表供选并记忆；
'3. 列出该表下所有字段及内容,并分页；
'由于不保证每个数据都有唯一不重复字段，暂无表的增改删计划
'未设密码，推荐文件夹下更名为任意文件名使用以保障数据安全

'问题反馈：15058593138 (同微信号)
'或发邮件：admin@ewuyi.net

'另外可能后续开发
'php access;php mysql;php csv;asp access;asp excel
'等版本,敬请关注

 dbcokie = Request.Cookies("dbname")
 dbbiaos = Request.Cookies("dbbiao")
 dbsname = request("db")
 rewname = request("tb")

if IsFile(dbsname)=True then

else
if IsFile(dbcokie)=True then
 dbsname = dbcokie
else
 dbsname = "down_zip.asa"
end if
end if

Function ShowFold(filepath)
Set fsoaaaa = Server.CreateObject("Scripting.FileSystemObject")
Set fileobj = fsoaaaa.GetFolder(server.mappath(filepath))
Set fsofolders = fileobj.SubFolders
For Each folder in fsofolders
foldername=folder.name
filepathes=filepath&""&foldername&"/"
ShowFold = ShowFold & ShowFold(filepathes)
Next
ShowFold = ShowFold & ShowFile(filepath)
End Function

Function ShowFile(filepath)
Set fsoeeee = Server.CreateObject("Scripting.FileSystemObject")
Set fileobj = fsoeeee.GetFolder(server.mappath(filepath))
Set fsofile = fileobj.Files
For Each file in fsofile 
filed = ""&filepath&""&file.name&""
if instr(file.name&"@",".mdb@")>0 or instr(file.name&"@",".asa@")>0 then
if filed = dbsname then
Response.Cookies("dbname")=dbsname
Response.Cookies("dbname").Expires=(now()+7) 
ShowFile = ShowFile& "<option value="""&filed&""" selected>"&filed&"</option>"&vbcrlf
else
ShowFile = ShowFile& "<option value="""&filed&""">"&filed&"</option>"&vbcrlf
end if
end if
Next
End Function

Function IsFile(FilePath)
 Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
 If (Fso.FileExists(Server.MapPath(FilePath))) Then
 IsFile=True
 Else
 IsFile=False
 End If
 Set Fso=Nothing
End Function

Function tabar(tabatxt)
tabar = "<table cellspacing=""0"" cellpadding=""0""><tr><td>"&tabatxt&"</td></tr></table>"
End Function

Function getpage(currentpage,totalpage) 
if currentpage mod 10 = 0 then 
Sp = currentpage \ 10 
else 
Sp = currentpage \ 10 + 1 
end if 
Pagestart = (Sp-1)*10+1 
Pageend = Sp*10 
strSplit = "<a href=""?p=1"&thekeys&""" title=""第1页"">第1页</a>" &vbcrlf
if Sp > 1 then strSplit = strSplit & "<a href=""?p="&Pagestart-10&""&thekeys&""">前10页</a>" &vbcrlf
for j=PageStart to Pageend 
if j > totalpage then exit for 
if "_"&j = "_"&currentpage then 
strSplit = strSplit & "<font color=""red"">第"&j&"页</font>" &vbcrlf
else 
strSplit = strSplit & "<a href=""?p="&j&""&thekeys&""" title=""第"&j&"页"">第"&j&"页</a>" &vbcrlf
end if 
next 
if Sp*10 < totalpage then strSplit = strSplit & "<a href=""?p="&Pagestart+10&""&thekeys&""">后10页</a>" &vbcrlf
strSplit = strSplit & "<a href=""?p="&totalpage&""&thekeys&""">最后1页</a>" &vbcrlf
strSplit = strSplit & "<!---"&currentpage&":"&totalpage&"--->" &vbcrlf
getpage = strSplit 
End Function


response.write "<div style=""margin:0 auto;overflow:auto;width:99%;height:95vh;"">"

response.write "<table cellspacing=""0"" cellpadding=""0""><tr>"
response.write "<td width=""100"">选择数据库:<br>"&vbcrlf
response.write "<select onchange=""window.location='?db='+this.value;"" />"&vbcrlf
 lister = ShowFold("/")
if len(lister)>10 then
response.write lister
else
response.write "<option value="""">无access数据文件</option>"&vbcrlf
end if
response.write "</select>"&vbcrlf
response.write "</td></tr></table>"&vbcrlf

if IsFile(dbsname)=True then
else
response.write tabar("暂未选择数据"&dbsname)
response.End()
end if
set conn=Server.CreateObject("ADODB.Connection")
conn.open "DRIVER=Driver do Microsoft Access (*.mdb);UID=admin;PWD=;DBQ="&Server.MapPath(dbsname)
If Err Then
err.Clear
response.write tabar("数据异常无法读取"&dbsname)
response.End()
End If

'列出所有表
Const adSchemaTables = 20
set objSchema = Conn.OpenSchema(adSchemaTables)
ia = 0
rows = "-"
Do While Not objSchema.EOF
	if objSchema("TABLE_TYPE") = "TABLE" then
ia = ia+1
rowname = objSchema("TABLE_NAME")
rows = rows&rowname&"-"
if rowname = rewname then
 rawhtml = rawhtml & "<option value="""&rowname&""" selected>"&rowname&"</option>"&vbcrlf
else
 rawhtml = rawhtml & "<option value="""&rowname&""" >"&rowname&"</option>"&vbcrlf
end if
end if
objSchema.MoveNext
Loop
if len(rewname)<1 then rewname = rowname
if ia < 1 then
 rawhtml = "<option value="""" selected>暂未发现数据表</option>"
end if
 rawhtml = "<select onchange=""window.location='?db="&dbsname&"&tb='+this.value;"" />"&rawhtml&"</select>"
objSchema.Close
set objSchema = nothing

response.write "<table cellspacing=""0"" cellpadding=""0"">"&vbcrlf
response.write "<tr><td>"&rawhtml&"</td></tr>"&vbcrlf
response.write "</table>"&vbcrlf

page=request("p")
if isnumeric(page)=false or len(page)=0 then
page="1"
end if
thekeys = "&db="&dbsname&"&tb="&rewname&""
if instr(rows,"-"&rewname&"-")>0 then
else
rewname = Request.Cookies("dbbiao")
if instr(rows,"-"&rewname&"-")<1 then
response.write tabar("请合理选择数据表哦"&rows)
response.End()
end if
end if

Response.Cookies("dbbiao")=rewname
Response.Cookies("dbbiao").Expires=(now()+7) 

Response.Write "<table cellspacing=""0"">"&vbcrlf
Response.Write "<caption align='center'>"&rewname&"表</caption>"&vbcrlf
set rsdo=Server.CreateObject("ADODB.RecordSet")
sqldo="select * from ["&rewname&"] "
rsdo.open sqldo,conn,1,1
rsdo.PageSize=10
lies = rsdo.fields.count
 tnames="---"
 response.write "<tr class=""tt"">"&vbcrlf
 for i = 0 to lies - 1 '循环字段名
 lieti = rsdo.fields.item(i).name
 response.write "<td><nobr>" & lieti & "</nobr></td>"&vbcrlf
 tnames=tnames&lieti&"---"
 next
 response.write "</tr>"&vbcrlf

If not (rsdo.bof and rsdo.eof) then
 rsdo.AbsolutePage=page
 for k=1 to rsdo.PageSize
 response.write "<tr>"&vbcrlf
 for i = 0 to lies - 1
 curValue = rsdo.fields.item(i).value
 If IsNull(curValue) or len(curValue)<1 Then
 curValue="&nbsp;"
 End If
 response.write "<td>" & curValue & "</td>"&vbcrlf
 next
 response.write "</tr>"&vbcrlf

 rsdo.movenext
 If rsdo.EOF Then Exit For
 next

rc=rsdo.RecordCount
ps=rsdo.PageSize
pc=rsdo.PageCount
if pc>1 then
response.write "<tr><td colspan="""&lies&""" class=""titi"">"
response.write getPage(page,pc)
response.write "</td></tr>"
end if
rsdo.close
set rsdo=nothing

else
response.write "<tr><td colspan="""&lies&""">"
response.write "<p>暂没查询信息！</p>"
response.write "</td></tr>"
end if
Response.Write "</table>"&vbcrlf
%></body>
</html>