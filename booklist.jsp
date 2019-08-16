<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
  <%--  <%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %> --%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>List of books</title>
</head>
<body>
<div align="center"><h1>Book Management System</h1></div>

<h2 align="center"><a href="/add">Add new book</a>  &nbsp;
<a href="/list">List all books</a></h2>

<div align="center">
<div align="center"><h2>Listof books</h2></div>
<table border="2 px solid black">
<tr>
<th>I.D</th>
<th>Title</th>
<th>Author</th>
<th>Price</th>
<th>Action</th>
<%-- <c:forEach var="book" items="${listbook}">
 <tr>
 <td><c:out value="${book.id}"></c:out></td>
 <td><c:out value="${book.title }"></c:out></td>
 <td><c:out value="${book.author }"></c:out></td>
 <td><c:out value="${book.price }"></c:out><td>
 <td><a href="/edit?<c:out value="${book.id}"">edit</a>&nbsp;&nbsp;
 <a href="/delete?<c:out value="${book.id}""></a></td>
 </tr>
</c:forEach> --%>
</tr>

</table>



</div>
</body>
</html>