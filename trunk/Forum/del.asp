<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->
<%sub del()
name=request.querystring("hostname")
code=request.Form("code")
id=request.querystring("id")
if code="" and id<>"" then%>
<FORM METHOD="post" ACTION="del.asp?id=<%=id%>" ID="Form1">
<%set rsname=conn.execute("select * from bbs where id=" & id)%>
       <br>
       将要删除&nbsp;<%=rsname("name")%>&nbsp;的大作: &nbsp;<%=rsname("title")%><br /><br />
       连同文章的回复都将被删除!<br><br />
       作者密码：<INPUT TYPE="password" name="code" SIZE="28" ID="Password1" style="font-size:12px">
       <INPUT TYPE="submit" VALUE="删除" ID="Submit1" NAME="Submit1"> 
</FORM>
<%elseif id="" then%>
还没有选择帖子,<a href="index.asp">返回首页</a>选择一个帖子删除
<%else
set rsname=conn.execute("select * from bbs where id=" & id)
set rs=conn.execute("select * from author where name='" &rsname("name")& "' and code= '" &code& "'")
      if not rs.eof then
       conn.execute("Delete from bbs where id="&id&" ")
       conn.execute("delete from bbbs where hostid='"&id&"'")%>
       <br>
       <center>删除成功！<a href="index.asp">查看帖子</a></center>
      <%else%>
       <br>
       用户名或密码错误!<a href= "del.asp?id=<%=id%>">返回</a>继续操作.<br>
       <a href="index.asp">返回首页</a>
      <%end if
set rs=nothing
set rsname=nothing
end if
end sub%>
<html><head>
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
<title>删除</title></head>
<body>
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
<table width="770" border="0" cellspacing="0" cellpadding="0" align="center"  class="top">
      <tr>
        <td id="tbul">&nbsp;</td>
        <td id= "tbum"><%headtext()%></td>
        <td id="tbur">&nbsp;</td>
      </tr>
      <tr>
        <td id="tbml">&nbsp;</td>
        <td id="tbmm"><%del()%></td>
        <td id="tbmr">&nbsp;</td>
      </tr>
  
      <tr>
        <td id="tbll">&nbsp;</td>
        <td id="tblm">&nbsp;</td>
        <td id="tblr">&nbsp;</td>
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

