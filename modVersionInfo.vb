Option Strict Off
Option Explicit On

Imports ESRI.ArcGIS

'Module modVersionInfo

Public Class ArcInit

    Public m_Version As Integer
    Public m_VerStr As String
    <Security.Permissions.EnvironmentPermissionAttribute(Security.Permissions.SecurityAction.LinkDemand, Unrestricted:=True)> _
    Public Function LoadVersionAndCheckOutLicense(ByRef sErr As String) As Boolean
        LoadVersionAndCheckOutLicense = True
        Dim VersMan As New ArcGISVersionLib.VersionManager
        Dim pVersion As ArcGISVersionLib.IArcGISVersion
        Dim bLoadOK As Boolean

        Dim sVer As String = vbNullString
        Dim fvVer As FileVersionInfo = GetArcGISVersion()
        If fvVer Is Nothing Then
            LoadVersionAndCheckOutLicense = False
            AddString(sErr, "Could not get ArcGIS version")
            Exit Function
        End If
        sVer = fvVer.FileMajorPart.ToString & "." & fvVer.FileMinorPart.ToString

        'register version
        If m_Version > 93 Then
            pVersion = VersMan
            'vista fix - need to do it twice!
            bLoadOK = pVersion.LoadVersion(ArcGISVersionLib.esriProductCode.esriArcGISDesktop, sVer)
            bLoadOK = pVersion.LoadVersion(ArcGISVersionLib.esriProductCode.esriArcGISDesktop, sVer)
            If Not bLoadOK Then
                LoadVersionAndCheckOutLicense = False
                AddString(sErr, "Load version failed")
                Exit Function
            End If
        End If

        'check license
        Dim pAoInitialize As esriSystem.IAoInitialize
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
            GetArcGISVersion = FileVersionInfo.GetVersionInfo(GetArcDir() & "bin\" & sDLLFile)
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
            '10.2
            GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}\Desktop10.2", esriKey), "InstallDir", Nothing)
            '10.1
            If Len(GetArcDir) < 1 Then GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}\Desktop10.1", esriKey), "InstallDir", Nothing)
            '10
            If Len(GetArcDir) < 1 Then GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}\Desktop10.0", esriKey), "InstallDir", Nothing)
            '9.3
            If Len(GetArcDir) < 1 Then GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}", esriKey), "InstallDir", Nothing)
            '9.2
            If Len(GetArcDir) < 1 Then GetArcDir = GetRegistryValue(String.Format("HKEY_LOCAL_MACHINE\Software\{0}\ArcInfo\Desktop\8.0", esriKey), "InstallDir", Nothing)
        Next
        If Len(GetArcDir) < 1 Then GetArcDir = "C:\ArcGIS\"
        If Right(GetArcDir, 1) <> "\" Then GetArcDir = GetArcDir & "\"

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
