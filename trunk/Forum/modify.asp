<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->
<%sub modify()
id=request.querystring("id")
set show=conn.execute("select*from bbs where id="&id&"")%>
<br>
<table border="0" cellpadding="0" cellspacing="0" 
        style="border-collapse: collapse; "
        width=600 align=center>
<form method="POST" action="modifysave.asp">
       <tr><td>
        <br>
        <input type="hidden" name="id" value=<%=show("id")%> size="20"><!--不要修改这个！ -->
        <input type="hidden" name="name" value=<%=show("name")%> size="20"><p>
        大名：<%=show("name")%>&nbsp;密码：<input type="password" name="code" size="20" style="font-size:12px"><br>
        标题：<input type="text" name="title" value=<%=show("title")%> size="76"><br>
        内容：<br>
        <textarea rows="11" name="body" cols="80"><%=show("body")%></textarea><br>
        <input type="submit" value="修改" name="B1"><input type="reset" value="重置" name="B2"><p>
       </td></tr>
</form>
</table>
<%set show=nothing
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
<title>修改</title></head>
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
		<table width="770" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td id="tbul">&nbsp;</td>
        <td id= "tbum"><%headtext()%></td>
        <td id="tbur">&nbsp;</td>
      </tr>
      <tr>
        <td id="tbml">&nbsp;</td>
        <td id="tbmm"><%modify()%></td>
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

