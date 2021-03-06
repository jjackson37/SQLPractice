﻿Imports System.Data.SqlClient

''' <summary>
''' Contains all the functionality related to the database
''' </summary>
Public Class Records

#Region "Var and Consts"
    ''' <summary>
    ''' The DBConnectionString must be edited to the correct file directory for the program to function
    ''' </summary>
    Private Const DBConnectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""C:\Users\jjackson\Documents\Visual Studio 2015\Projects\SQLPractice\SQLPractice\TestDB.mdf"";Integrated Security=True"
    ''' <summary>
    ''' Shows exceptions and messages for DB connections
    ''' </summary>
    Private Const Debug As Boolean = True
    Dim connection As SqlConnection
#End Region

#Region "Methods"

    ''' <summary>
    ''' Fills connection with the DBConnectionString
    ''' </summary>
    Private Sub InitializeConnection()
        connection = New SqlConnection(DBConnectionString)
    End Sub

    ''' <summary>
    ''' Opens the DB connection
    ''' </summary>
    Private Sub DBOpen()
        Try
            InitializeConnection()
            connection.Open()
            If Debug Then
                Console.WriteLine("+DB Connection opened successfully")
            End If
        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine(ex.Message)
            Console.WriteLine("Press any key to continue...")
            Console.ReadKey()
        End Try
    End Sub

    ''' <summary>
    ''' Closes the DB connection
    ''' </summary>
    Private Sub DBClose()
        Try
            connection.Close()
            connection.Dispose()
            If Debug Then
                Console.WriteLine("-DB Connection closed successfully")
            End If
            Console.WriteLine("Press any key to continue...")
            Console.ReadKey()
        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine(ex.Message)
            Console.WriteLine("Press any key to continue...")
            Console.ReadKey()
        End Try
    End Sub

    ''' <summary>
    ''' Prints all records in the DB
    ''' </summary>
    Public Sub PrintRecords()
        Try
            Console.WriteLine("Order by [i]d, [f]orename, [s]urname or [d]ate of birth?")
            Dim sortByCol As String = Console.ReadKey.KeyChar
            DBOpen()
            Dim command As SqlCommand
            Select Case sortByCol
                Case "f"
                    command = New SqlCommand("SELECT * FROM TestTable ORDER BY ForeName", connection)
                Case "s"
                    command = New SqlCommand("SELECT * FROM TestTable ORDER BY LastName", connection)
                Case "d"
                    command = New SqlCommand("SELECT * FROM TestTable ORDER BY DoB", connection)
                Case Else
                    command = New SqlCommand("SELECT * FROM TestTable", connection)
            End Select
            Dim reader As SqlDataReader = command.ExecuteReader
            Console.WriteLine("ID Forename       Surname        DoB")
            While reader.Read
                Console.Write(reader(0))
                Console.Write("  ")
                Console.Write(reader(1))
                Console.Write(reader(2))
                Console.WriteLine(CDate(reader(3)).ToShortDateString)
            End While
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    ''' <summary>
    ''' Adds a new record into the DB
    ''' </summary>
    Public Sub AddRecord()
        Dim foreName, lastName, DoB As String
        Try
            Console.WriteLine("Enter Forename:")
            foreName = Console.ReadLine()
            If ValidateUserInput(foreName) = False Then
                Throw New Exception("Invalid Forename input")
            End If
            Console.WriteLine("Enter Lastname:")
            lastName = Console.ReadLine()
            If ValidateUserInput(lastName) = False Then
                Throw New Exception("Invalid Surname input")
            End If
            Console.WriteLine("Enter DoB (dd/mm/yyyy):")
            DoB = Console.ReadLine()
            Dim format As String = "INSERT INTO TestTable( ForeName, LastName, DoB) VALUES( {0}{1}{0}, {0}{2}{0}, {0}{3}{0})"
            Dim rowInput As String = String.Format(format, Chr(39), foreName, lastName, DoB)
            DBOpen()
            Dim Command As SqlCommand = New SqlCommand(rowInput, connection)
            Command.BeginExecuteNonQuery()
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    ''' <summary>
    ''' Searchs for a record in the DB based on a column
    ''' </summary>
    Public Sub SearchRecords()
        Try
            Dim results As New ArrayList()
            Dim colName, commandString, searchTerm As String
            Console.WriteLine("Search using [f]orename or [s]urname?")
            Dim userInput As String = Console.ReadKey.KeyChar
            Console.Beep()
            Console.WriteLine()
            Select Case userInput
                Case "f"
                    colName = "ForeName"
                Case "s"
                    colName = "LastName"
                Case Else
                    Throw New Exception("Invalid selection")
            End Select
            Console.WriteLine("Enter the " & colName)
            searchTerm = Console.ReadLine()
            If ValidateUserInput(searchTerm) = False Then
                Throw New Exception("Invalid input")
            End If
            Dim format As String = "SELECT Id, ForeName, LastName, DoB FROM TestTable WHERE {1} = {0}{2}{0}"
            DBOpen()
            commandString = String.Format(format, Chr(39), colName, searchTerm)
            Dim command As SqlCommand = New SqlCommand(commandString, connection)
            Dim reader As SqlDataReader = command.ExecuteReader
            While reader.Read
                Console.Write(reader(0).ToString & " ")
                Console.Write(reader(1).ToString & " ")
                Console.Write(reader(2).ToString & " ")
                Console.WriteLine(CDate(reader(3)).ToShortDateString)
                results.Add(reader(0))
            End While
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    ''' <summary>
    ''' Deletes a record from the DB
    ''' </summary>
    Public Sub DeleteRecord()
        Try
            Console.WriteLine("Enter the ID of the record you want to delete:")
            Dim idInput As Integer = CInt(Console.ReadLine())
            If ValidateUserInput(idInput.ToString) = False Then
                Throw New Exception("Invalid ID input")
            End If
            Dim deleteQuery As String = "DELETE FROM TestTable WHERE Id = " & idInput
            DBOpen()
            Dim command As SqlCommand = New SqlCommand(deleteQuery, connection)
            command.ExecuteNonQuery()
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    ''' <summary>
    ''' Edits a record from the DB
    ''' </summary>
    Public Sub EditRecord()
        Try
            Dim queryString, editInput, colEdit As String
            Console.WriteLine("Enter the ID of the record you want to edit:")
            Dim idInput As Integer = CInt(Console.ReadLine())
            If ValidateUserInput(idInput.ToString) = False Then
                Throw New Exception("Invalid ID input")
            End If
            Dim format As String = "SELECT ForeName, LastName, DoB FROM TestTable WHERE Id = '{0}'"
            queryString = String.Format(format, idInput)
            DBOpen()
            Dim command As SqlCommand = New SqlCommand(queryString, connection)
            Dim reader As SqlDataReader = command.ExecuteReader
            Console.WriteLine("Forename:      Surname:       DoB:")
            While reader.Read
                Console.WriteLine(reader(0).ToString & reader(1).ToString & reader(2).ToString)
            End While
            DBClose()
            Console.WriteLine("Select what you would like to edit")
            Console.WriteLine("[f]orename, [s]urname or [d]ob")
            Dim userInput As String = Console.ReadKey.KeyChar
            Select Case userInput
                Case "f"
                    colEdit = "ForeName"
                Case "s"
                    colEdit = "LastName"
                Case "d"
                    colEdit = "DoB"
                Case Else
                    Throw New Exception("Invalid selection")
            End Select
            Console.WriteLine()
            Console.WriteLine("Enter the new value:")
            editInput = Console.ReadLine()
            If colEdit <> "DoB" Then
                If ValidateUserInput(editInput) = False Then
                    Throw New Exception("Invalid input")
                End If
            End If
            format = "UPDATE TestTable SET {1}='{2}' WHERE Id={0}"
            queryString = String.Format(format, idInput, colEdit, editInput)
            DBOpen()
            Dim command2 As SqlCommand = New SqlCommand(queryString, connection)
            command2.ExecuteNonQuery()
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    ''' <summary>
    ''' Validates a user input and returns true if valid and false if not
    ''' </summary>
    ''' <param name="UserInput">String to validate</param>
    ''' <returns>Boolean equal to if the UserInput passes the validation</returns>
    Private Function ValidateUserInput(userInput As String) As Boolean
        Dim valid As Boolean = True
        If userInput.Contains(" ") Then
            valid = False
        ElseIf userInput = "" Then
            valid = False
        End If
        Return valid
    End Function

#End Region
End Class