<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body><%dim code,strsql
code=request.QueryString("code")
set a=server.CreateObject("adodb.connection")
a.open"driver=sql server;server=pc07;uid=sa;pwd=;database=shuju"
strsql="delete shuju where code='"&code&"'"
a.execute strsql
a.close
response.Redirect("shujuku2.asp")


%>
</body>

</html>
