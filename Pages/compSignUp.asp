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
    tablePtr.Open "Company_main_info", connection, 1, 3

'Getting started with adding data to the table:
    tablePtr.AddNew 

'Passing data to fields in a table
    tablePtr("company_name") =  request.form("companyName")
    tablePtr("company_kind_business") =  request("kindOfBusiness")
    tablePtr("company_city") =  request("companyCity")
    tablePtr("kind_of_company") =  request("kindOfCompany")
    tablePtr("company_admin_name") =  request("adminName")
    tablePtr("company_phone") =  request("companyPhone")
    tablePtr("company_admin_password") =  request("companyPass")
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