'ReadMxd - export map document properties to a text file
'Copyright (C) 2015 Jon Morris

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.

Option Strict Off
Option Explicit On

Imports ESRI.ArcGIS

'Module modVersionInfo

Public Class ArcInit

    Public m_Version As Integer
    Public m_VerStr As String
    Protected pAoInitialize As esriSystem.IAoInitialize

    Public Sub Shutdown()
        If Not pAoInitialize Is Nothing Then
            pAoInitialize.Shutdown()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pAoInitialize)
            pAoInitialize = Nothing
        End If
    End Sub

    <Security.Permissions.EnvironmentPermissionAttribute(Security.Permissions.SecurityAction.LinkDemand, Unrestricted:=True)> _
    Public Function LoadVersionAndCheckOutLicense(ByRef sErr As String) As Boolean
        LoadVersionAndCheckOutLicense = True

        Dim sVer As String = vbNullString
        Dim fvVer As FileVersionInfo = GetArcGISVersion()
        If fvVer Is Nothing Then
            LoadVersionAndCheckOutLicense = False
            AddString(sErr, "Could not get ArcGIS version")
            Exit Function
        End If
        sVer = fvVer.FileMajorPart.ToString & "." & fvVer.FileMinorPart.ToString

        'register version
        If Not ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.Desktop) Then
            LoadVersionAndCheckOutLicense = False
            AddString(sErr, "Load version failed")
            Exit Function
        End If

        'check license
        pAoInitialize = New esriSystem.AoInitialize

        Dim licenseStatus As esriSystem.esriLicenseStatus
        Dim productCode As esriSystem.esriLicenseProductCode
        Dim strSoftwareClass As String
        strSoftwareClass = GetRegistryValue("HKEY_LOCAL_MACHINE\Software\ESRI\License" & sVer, "SOFTWARE_CLASS", Nothing)
        If Len(strSoftwareClass) < 1 Then _
            strSoftwareClass = GetRegistryValue("HKEY_LOCAL_MACHINE\Software\ESRI\License", "SOFTWARE_CLASS", Nothing)

        Select Case LCase(strSoftwareClass)
            Case "professional"
                productCode = esriSystem.esriLicenseProductCode.esriLicenseProductCodeAdvanced
            Case "editor"
                productCode = esriSystem.esriLicenseProductCode.esriLicenseProductCodeStandard
            Case "viewer"
                productCode = esriSystem.esriLicenseProductCode.esriLicenseProductCodeBasic
            Case Else
                productCode = esriSystem.esriLicenseProductCode.esriLicenseProductCodeAdvanced
        End Select

        'initialize license
        licenseStatus = pAoInitialize.Initialize(productCode)
        If Not (licenseStatus = esriSystem.esriLicenseStatus.esriLicenseCheckedOut) Then
            AddString(sErr, "Error: License Initialization Failed = " & GetLicenseStatus(licenseStatus))
            LoadVersionAndCheckOutLicense = False
        End If
        'check out a maplex extension license
        licenseStatus = pAoInitialize.IsExtensionCodeAvailable(productCode, esriSystem.esriLicenseExtensionCode.esriLicenseExtensionCodeMLE)
        If licenseStatus = esriSystem.esriLicenseStatus.esriLicenseAvailable Then
            licenseStatus = pAoInitialize.CheckOutExtension(esriSystem.esriLicenseExtensionCode.esriLicenseExtensionCodeMLE)
            If Not (licenseStatus = esriSystem.esriLicenseStatus.esriLicenseCheckedOut) Then
                AddString(sErr, "Maplex license did not check out. Status = " & GetLicenseStatus(licenseStatus))
                LoadVersionAndCheckOutLicense = False
            End If
        End If
    End Function

    Function GetRegistryValue(ByVal KeyName As String, ByVal ValueName As String, ByRef DefaultValue As Object) As Object

        GetRegistryValue = Microsoft.Win32.Registry.GetValue(KeyName, ValueName, DefaultValue)

    End Function

    <Security.Permissions.EnvironmentPermissionAttribute(Security.Permissions.SecurityAction.LinkDemand, Unrestricted:=True)> _
    Public Function GetArcGISVersion(Optional ByRef sDLLFile As String = "AfCore.dll") As FileVersionInfo

        Try
            GetArcGISVersion = FileVersionInfo.GetVersionInfo(Path.Combine(GetArcDir(), "bin", sDLLFile))
            m_Version = GetArcGISVersion.FileMajorPart * 10 + GetArcGISVersion.FileMinorPart
            m_VerStr = GetArcGISVersion.FileVersion.ToString
        Catch ex As FileNotFoundException
            Return Nothing
        End Try

    End Function

    'return ArcGIS install directory
    Function GetArcDir() As String

        GetArcDir = ""
        'have a few goes at getting install dir, depending on version. use c:\arcgis if can't find it
        For Each esriKey As String In {"Wow6432Node\ESRI", "ESRI"}
            '10.x
            For Each esriVersion As String In {"10.4", "10.3", "10.2", "10.1", "10.0"}
                GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}\Desktop{1}", esriKey, esriVersion), "InstallDir", Nothing)
                If Len(GetArcDir) > 1 Then Return GetArcDir
            Next
            '9.3
            GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}", esriKey), "InstallDir", Nothing)
            If Len(GetArcDir) > 1 Then Return GetArcDir
            '9.2
            GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}\ArcInfo\Desktop\8.0", esriKey), "InstallDir", Nothing)
            If Len(GetArcDir) > 1 Then Return GetArcDir
        Next
        Return "C:\ArcGIS\"

    End Function

    Function GetLicenseStatus(ByRef status As esriSystem.esriLicenseStatus) As String
        Select Case status
            Case esriSystem.esriLicenseStatus.esriLicenseAlreadyInitialized
                GetLicenseStatus = "Already Initialized"
            Case esriSystem.esriLicenseStatus.esriLicenseAvailable
                GetLicenseStatus = "Available"
            Case esriSystem.esriLicenseStatus.esriLicenseNotLicensed
                GetLicenseStatus = "Not Licensed"
            Case esriSystem.esriLicenseStatus.esriLicenseCheckedOut
                GetLicenseStatus = "Checked Out"
            Case esriSystem.esriLicenseStatus.esriLicenseCheckedIn
                GetLicenseStatus = "Checked In"
            Case esriSystem.esriLicenseStatus.esriLicenseFailure
                GetLicenseStatus = "Failure"
            Case esriSystem.esriLicenseStatus.esriLicenseNotInitialized
                GetLicenseStatus = "Not Initialized"
            Case Else
                GetLicenseStatus = "Unknown Status"
        End Select
    End Function

    Sub AddString(ByRef str As String, ByVal msg As String)
        If str Is Nothing Then
            str = msg
        ElseIf str.Length < 1 Then
            str = msg
        Else
            str = str & vbNewLine & msg
        End If
    End Sub
    'End Module
End Class
