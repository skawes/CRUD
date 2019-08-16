<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Add new book</title>
</head>
<body>
<div align="center"><h1>Book Management System</h1></div>

<h2 align="center"><a href="/add.jsp">Add new book</a>  &nbsp;
<a href="/list">List all books</a></h2>
<h3 align="center">Add new book</h3>
<form action="/addbook">
<table border="2 px solid black" align="center">
<tr>
<th>Title</th><td><input type="text"> </td>
</tr>
<tr>
<th>Author</th><td><input type="text">  </td>
</tr>
<tr>
<th>Price</th><td><input type="text"> </td>
</tr>
</table>
<br>
<div align="center"><input " type="submit" value="add" ></div>
</form>

</body>
</html>