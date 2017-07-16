<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>

<body>
<%dim xingming,sex,age,address,tel,code,sex1
xingming=request.Form("xingming")
sex=request.Form("sex")
age=request.Form("age")
address=request.Form("address")
tel=request.Form("tel")
code=request.Form("code")
if sex=("0") then
sex1="男"
else
sex1="女"
end if



%><table width="742" border="1">
    <tr align="center">
      <td colspan="2">您的注册信息为</td>
    </tr>
    <tr>
      <td width="151" align="center">姓名</td>
      <td width="575"><%response.Write(xingming)%></td>
    </tr>
    <tr>
      <td align="center">性别</td>
      <td><p><%response.Write(sex1)%>
       
        <br />
      </p></td>
    </tr>
    <tr>
      <td align="center">年龄</td>
      <td><%response.Write(age)%></td>
    </tr>
    <tr>
      <td align="center">住址</td>
      <td><%response.Write(address)%></td>
    </tr>
    <tr>
      <td align="center">电话</td>
      <td><%response.Write(tel)%>;</td>
    </tr>
    <tr>
      <td align="center">学号</td>
      <td><%response.Write(code)%></td>
    </tr>
  </table>
<p>&nbsp;</p>
<p><a href="../index.html">返回主页
</a></p>
</body>
</html>
