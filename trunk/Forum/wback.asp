<!--#include file="conn.asp"--> 
<link href="main.css" rel="stylesheet" type="text/css" />
<%
idh=Replace(Request.Form("id_h"),"'","''")
nameb=Replace(Request.Form("name_b"),"'","''") 
titleb=Replace(Request.Form("title_b"),"'","''") 
bodyb=Replace(Request.Form("body_b"),"'","''") 
codeb=Replace(Request.Form("code_b"),"'","''")
set rs=conn.execute("select * from author where name='" &nameb& "' and code='" &codeb& "'")
sql="select * from author where name='" &nameb& "' and code='" &codeb& "'"
if nameb="" or titleb="" or bodyb="" or idh="" or rs.eof then%><title>回复保存</title>
有数据为空/错误的用户名或密码!<a href="show.asp?id=<%=idh%>">查看帖子</a>,点此<a href="register.asp">注册</a>
<%else
sql="update bbs set countwb=countwb+1 where id=" &idh
conn.execute(sql)
sql="insert into bbbs(hostid,btime,bname,bcontent,btitle) " & _
         "values('"& idh & "','" & now() &"','"& nameb &"','"& bodyb &"','"& titleb &"')"
conn.execute(sql)%> 
发表成功！<a href="index.asp">回到首页</a> <br>
<a href="show.asp?id=<%=idh%>">查看帖子</a>
<%response.Redirect("show.asp?id="&idh)
end if %>
