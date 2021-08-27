<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<% 


'--------------
'Making the VT connection:
    Set connection = CreateObject("ADODB.Connection") 
'VT opening:
    connection.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ="& Server.MapPath("../stajProjectDB.mdb"))
'Creating the table object:
    Set tablePtr = server. CreateObject("ADODB.Recordset")
'Opening the table:
    tablePtr.Open "Users_main_info", connection, 1, 3

'Getting started with adding data to the table:
    tablePtr.AddNew 

'Passing data to fields in a table
    tablePtr("student_first_name") =  request.form("firstName")
    tablePtr("student_last_name") =  request("lastName")
    tablePtr("student_date") =  request("birthDay")
    tablePtr("student_id") =  request("studentId")
    tablePtr("student_bolum") =  request("bulum")
    tablePtr("student_city") =  request("city")
    tablePtr("student_password") =  request("password")

'aktarma islemi birince tablonun guncellenmesi:
    tablePtr.Update

'tablonun kapatilmasi:
    tablePtr.close
    set tablePtr= Nothing
'baglantinin kesilmesi:
    connection.close
    set connection= Nothing

response.write "Data Entry Has Been Made"
%>
<p><a href="../index.html">back to Homepage</a></p>