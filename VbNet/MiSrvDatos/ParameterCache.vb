Option Explicit On
Option Strict On
Imports System.IO
Imports System
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading



''' <summary>
''' This class is used to cache parameters based 
''' on the connectionstring and procedurename
''' </summary>
Friend Class ParameterCache
    ''' <summary>
    ''' A synchronized hashtable used to cache the parameters.
    ''' </summary>
    Private Shared ReadOnly mHashTable As Hashtable = Hashtable.Synchronized(New Hashtable())

    ''' <summary>
    ''' Default constructor
    ''' </summary>
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Gets whether the given command has a cached parameter set 
    ''' </summary>
    ''' <param name="command">The command to check for cached parameters</param>
    ''' <returns>True if the command exists in the cache, otherwise it return false</returns>
    Friend Shared Function IsParametersCached(command As SqlCommand) As Boolean
        Dim key As String = command.Connection.ConnectionString + ":" + command.CommandText
        Return mHashTable.Contains(key)
    End Function


    ''' <summary>
    ''' Adds the given command's parameters to the parameter cache 
    ''' </summary>
    ''' <param name="command">The command that holds the parameters to be cached</param>
    Friend Shared Sub CacheParameters(command As SqlCommand)
        Dim originalParameters As SqlParameter() = New SqlParameter(command.Parameters.Count - 1) {}
        command.Parameters.CopyTo(originalParameters, 0)
        Dim parameters As SqlParameter() = CloneParameters(originalParameters)
        mHashTable(command.Connection.ConnectionString + ":" + command.CommandText) = parameters
    End Sub

    ''' <summary>
    ''' Gets a array of IDataParameter for the given command
    ''' </summary>
    ''' <param name="command">The command to get cached parameters</param>
    ''' <returns>An array of IDataParameter</returns>
    Friend Shared Function GetCachedParameters(command As SqlCommand) As SqlParameter()
        Dim originalParameters As SqlParameter() = DirectCast(mHashTable(command.Connection.ConnectionString + ":" + command.CommandText), SqlParameter())
        Return CloneParameters(originalParameters)
    End Function


    ''' <summary>
    ''' Used to create a copy of an array of IDataParameter
    ''' </summary>
    ''' <param name="originalParameters">The array of IDataParameter we want to copy</param>
    ''' <returns>An array of IDataParameter</returns>
    Private Shared Function CloneParameters(originalParameters As SqlParameter()) As SqlParameter()
        Dim clonedParameters As SqlParameter() = New SqlParameter(originalParameters.Length - 1) {}

        Dim i As Integer = 0, j As Integer = originalParameters.Length
        While i < j
            clonedParameters(i) = DirectCast(DirectCast(originalParameters(i), ICloneable).Clone(), SqlParameter)
            i += 1
        End While

        Return clonedParameters
    End Function
End Class





'Public Class ParameterCache
'    Implements IParameterCache
'    Private dictionary As New Dictionary(Of String, DbParameter())()
'    Private [syncLock] As New Object()

'    Public Function ContainsParameters(connectionString As String, storedProcedureName As String) As Boolean
'        Return dictionary.ContainsKey(GetCacheKey(connectionString, storedProcedureName))
'    End Function

'    Public Function GetParameters(connectionString As String, storedProcedure As String) As DbParameter()
'        Return CopyParameterArray(dictionary(GetCacheKey(connectionString, storedProcedure)))
'    End Function

'    Public Sub AddParameters(connectionString As String, storedProcedureName As String, parameterCollection As DbParameterCollection)
'        If Not ContainsParameters(connectionString, storedProcedureName) Then
'            SyncLock [syncLock]
'                If Not ContainsParameters(connectionString, storedProcedureName) Then
'                    Dim cacheKey As String = GetCacheKey(connectionString, storedProcedureName)
'                    Dim parameters As DbParameter() = GetParameterArray(parameterCollection)

'                    dictionary.Add(cacheKey, parameters)
'                End If
'            End SyncLock
'        End If
'    End Sub
'End Class

