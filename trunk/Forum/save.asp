<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->
<%
sub save()
name=Replace(Request.Form("name"),"'","''") 
title=Replace(Request.Form("title"),"'","''") 
body=Replace(Request.Form("body"),"'","''") 
code=Replace(Request.Form("code"),"'","''")
set savebbs=conn.execute("select * from author where name='" &name& "' " & "and code='" &code& "'" )
if name="" or title="" or body="" or code="" or savebbs.eof then
%> 
<br>
请<a href="index.asp">后退</a>填写完整资料/填写正确用户名和密码，你才能发表帖子！<br><br>
点此<a href="register.asp">注册</a>
<%
else
sql="insert into bbs(name,title,body,wtime,countwb)values('"& name & "','"&title&"','"&body& "','" & now() & "', 0)"
set savebbs=conn.execute(sql)
set savebbs=nothing
set savebbs=conn.execute("select * from bbs where name='" & name & "' order by wtime desc")
%> 
<br>
发表成功！<a href="show.asp?id=<%=savebbs("id")%>">查看帖子</a> |<a href="index.asp">返回论坛</a>

<%end if
set savebbs=nothing
end sub
%>
<html>
<head>
<link href="main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
	.top{
	vertical-align:top;}
	.btn{
	background-color:#CCFFCC;
	width:130px;
	height:30px;
	}

</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>保存</title>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (sub01改.psd) -->
<table id="__01" width="950" height="793" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="3">
			<img src="images/sub01改_01.gif" width="814" height="90" alt=""></td>
		<td colspan="2">
			<img src="images/sub01改_02.gif" width="136" height="90" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/sub01_03.gif" width="950" height="68" alt=""></td>
	</tr>
	<tr>
		<td><a href="index.asp"><img src="images/left_menu01.gif" width="219" height="29" alt="" border="0"></a></td>
		<td colspan="4" rowspan="2">
			<img src="images/sub01改_05.gif" width="731" height="49" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/left_menu02.gif" width="219" height="20" alt=""></td>
	</tr>
	<tr>
		<td rowspan="2" class="top">
			<img src="images/sub01改_07.gif" width="219" height="505" alt=""></td>
		<td rowspan="2">
			<img src="images/sub01改_08.gif" width="21" height="505" alt=""></td>
		<td colspan="2" background="images/txt.gif">
		<table width="770" border="0" cellspacing="0" align="center" ID="Table3">
      <tr>
        <td      id="tbul">　</td>
        <td id= "tbum"><%headtext()%></td>
        <td      id="tbur">　</td>
      </tr>
      <tr>
        <td id="tbml">　</td>
        <td id="tbmm"><%save()%></td>
        <td id="tbmr">　</td>
      </tr>
  
      <tr>
        <td id="tbll">　</td>
        <td id="tblm">　</td>
        <td id="tblr">　</td>
      </tr>
</table>
<!--#include file="footer.htm"--></td>
		<td rowspan="2">
			<img src="images/sub01改_10.gif" width="18" height="505" alt=""></td>
	</tr>
	<tr>
		<td colspan="2">
			<img src="images/sub01改_11.gif" width="692" height="14" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/sub01改_12.gif" width="950" height="80" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/分隔符.gif" width="219" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="21" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="574" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="118" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="18" height="1" alt=""></td>
	</tr>
</table>
<!-- End ImageReady Slices -->
</body>
</html>
