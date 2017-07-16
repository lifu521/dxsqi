<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body><%dim code,strsql,sex
code=request.QueryString("code")
set a=server.CreateObject("adodb.connection")
a.open"driver=sql server;server=pc07;uid=sa;pwd=;database=shuju"
set b=server.CreateObject("adodb.recordset")
b.open "select * from shuju where code= '"&code&"'",a,3,3

sex=b("sex")
if sex="true" then
 	sex="男"
else
	sex="女"
end if

%>
<table width="962" border="1">
  <tr>
    <td colspan="2">个人信息</td>
  </tr>
  <tr>
    <td width="419">学号</td>
    <td width="527"><%=b("code")%></td>
  </tr>
  <tr>
    <td>姓名</td>
    <td><%=b("xingming")%></td>
  </tr>
  <tr>
    <td>性别</td>
    <td><%=sex)%></td>
  </tr>
  <tr>
    <td>地址</td>
    <td><%=b("address")%></td>
  </tr>
  <tr>
    <td>电话</td>
    <td><%=b("tel")%></td>
  </tr>
  <tr>
    <td height="17">年龄</td>
    <td><%=b("age")%></td>
  </tr>
</table>

<p><a href="../index.html">返回首页</a></p>
</body>
</html>
