Option Strict Off
Option Explicit On

Imports ArcGISVersionLib
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports Ionic.Utils.Zip
Imports System.Linq
Imports System.Runtime.InteropServices



Module modReadMxd

    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)> _
    Public Function GetDesktopWindow() As IntPtr
    End Function

    Public sw As StreamWriter
    Public m_Version As Integer
    Public bAllLayers As Boolean
    Public bShowFullExp As Boolean
    Public bLyrFile As Boolean
    Public bReadSymbols As Boolean
    Public bReadLabels As Boolean
    Private pMapDocument As IMapDocument = Nothing
    Private pActiveView As IActiveView = Nothing
    Public mxdProps As clsMxdProps = Nothing
    Public Const ARRAY_SIZE As Long = 128
    Private Const sSummaryXls As String = "C:\temp\MxdSummary.xlsx"

    Public Sub Main()
        Dim i As Integer

        If Environment.GetCommandLineArgs().Length < 1 Then Exit Sub
        Dim sArgs() As String
        Dim sMxdName As String
        Dim bLocalLog, bExcel As Boolean
        sArgs = Environment.GetCommandLineArgs()
        sMxdName = vbNullString
        bExcel = False
        bAllLayers = False
        bLocalLog = False
        bShowFullExp = False
        bReadSymbols = False
        bReadLabels = False
        For i = 1 To UBound(sArgs)
            Select Case LCase(sArgs(i))
                Case "-a"
                    bAllLayers = True
                Case "-l"
                    bLocalLog = True
                Case "-x"
                    bExcel = True
                Case "-e"
                    bShowFullExp = True
                Case "-s"
                    bReadSymbols = True
                Case "-b"  'sorry, l is already taken!
                    bReadLabels = True
                Case Else
                    'reconstruct if spaces in file name
                    sMxdName = sMxdName & sArgs(i) & " "
            End Select
        Next

        'remove end space and any quotes
        sMxdName = Trim(Replace(sMxdName, Chr(34), ""))

        'add current path if none given
        If InStr(sMxdName, "\") = 0 Or InStr(sMxdName, "\") > 3 Then sMxdName = CurDir() & "\" & sMxdName

        'detect file type
        Select Case IO.Path.GetExtension(sMxdName).ToLower()
            Case ".txt"
                'read file list
                Dim sr As StreamReader = File.OpenText(sMxdName)
                Dim sLine As String = sr.ReadLine
                While Not sLine = vbNullString
                    OpenLogAndMxd(sLine, bLocalLog, bExcel)
                    sLine = sr.ReadLine
                End While
                sr.Close()
            Case ".mxd", ".msd", ".aprx"
                OpenLogAndMxd(sMxdName, bLocalLog, bExcel)
            Case ".lyr"
                bLyrFile = True
                OpenLogAndMxd(sMxdName, bLocalLog, bExcel)
        End Select

    End Sub

    Private Sub OpenLogAndMxd(ByVal sMxdName As String, ByVal bLocalLog As Boolean, _
                              ByVal bExcel As Boolean)
        Dim sLogName, sTmp, sType As String
        sType = IO.Path.GetExtension(sMxdName)

        'set log name
        If bLocalLog Then
            If sType.Equals(".mxd", StringComparison.CurrentCultureIgnoreCase) Then
                sLogName = Replace(sMxdName, sType, "_props.log")
            Else
                sLogName = Replace(sMxdName, ".", "_") + "_props.log"
            End If
        ElseIf StrComp(My.Application.Info.DirectoryPath, Environment.CurrentDirectory()) = 0 Then
            sLogName = My.Application.Info.DirectoryPath & "\MxdProps.log"
        Else
            If sType.Equals(".mxd", StringComparison.CurrentCultureIgnoreCase) Then
                sLogName = Environment.CurrentDirectory() & Replace(Mid(sMxdName, InStrRev(sMxdName, "\")), _
                                                                    sType, "_props.log")
            Else
                sLogName = Environment.CurrentDirectory() & Replace(Mid(sMxdName, InStrRev(sMxdName, "\")), ".", "_") + "_props.log"
            End If
        End If

        'check for ? in mxdname (replace with _)
        sLogName = IO.Path.GetInvalidPathChars.Aggregate(sLogName, Function(current, ch) Replace(current, ch, "_"))
        sLogName = Replace(sLogName, "?", "_")
        If Not Directory.Exists(IO.Path.GetDirectoryName(sLogName)) Then
            Console.WriteLine("Directory not found: " & sLogName)
            Exit Sub
        End If

        'open log file
        sw = File.CreateText(sLogName)
        sw.WriteLine(sMxdName & " Properties")
        sTmp = New String("=", Len(sMxdName) + 11)
        sw.WriteLine(sTmp)
        sw.WriteLine("ReadMxd version: " & My.Application.Info.Version.Major & _
                     "." & My.Application.Info.Version.Minor & _
                     "." & My.Application.Info.Version.Build & _
                     "." & My.Application.Info.Version.Revision)

        If Not File.Exists(sMxdName) Then
            sw.WriteLine(sMxdName & " not found")
        Else
            If sType.Equals(".msd", StringComparison.CurrentCultureIgnoreCase) Or _
                sType.Equals(".aprx", StringComparison.CurrentCultureIgnoreCase) Then
                ReadXMLArchive(sMxdName, bExcel)
            Else
                ReadProps(sMxdName, bExcel)
            End If
        End If

        'dispose global variables
        If Not pMapDocument Is Nothing Then
            If Not pActiveView Is Nothing Then _
                If pActiveView.IsActive Then pActiveView.Deactivate()
            pMapDocument.Close()
            releaseObject(pActiveView)
            releaseObject(pMapDocument)
        End If

        sw.Close()
    End Sub

    Private Sub ReadProps(ByVal sMxdName As String, ByVal bExcel As Boolean)
        Dim i As Integer
        Dim j As Integer
        Dim sConnType As String, sTmp As String

        'find arcgis version
        Dim arc As ArcInit = New ArcInit()
        Dim sError As String = vbNullString
        If arc.LoadVersionAndCheckOutLicense(sError) Then
            sw.WriteLine("ArcGIS version: " & arc.m_VerStr)
            m_Version = arc.m_Version
        Else
            sw.WriteLine(sError)
            Exit Sub
        End If

        'create new set of props
        mxdProps = New clsMxdProps
        sw.Flush()

        'layer file - go straight to layer props
        If bLyrFile Then
            'Create a GxLayer.
            Dim gxLayerCls As IGxLayer = New GxLayerClass
            Dim gxFile As IGxFile = gxLayerCls 'Implicit Cast.
            'Set the path for where the layer file is located on disk.
            gxFile.Path = sMxdName

            'file attributes
            sw.WriteLine("Last Modified: " & File.GetLastWriteTime(sMxdName))
            sw.WriteLine(vbCrLf & InsertTabs(1) & "Layer properties:")
            If Not gxLayerCls.Layer Is Nothing Then
                GetLayerProps(gxLayerCls.Layer(), 2)
            End If
            sw.WriteLine("")
            If bReadLabels Then WriteLabelSummary()
            If bReadSymbols Then WriteSymbolSummary()
            WriteGeneralSummary()
            sw.WriteLine(vbCrLf & "OK")
            Exit Sub
        End If

        If bAllLayers Then sw.WriteLine("Showing all layers")
        If bReadSymbols Then sw.WriteLine("Reading symbols")
        If bReadLabels Then sw.WriteLine("Reading labels")
        'open mxd
        Try
            pMapDocument = New MapDocument
            pMapDocument.Open(sMxdName)

            pActiveView = pMapDocument.ActiveView
            pActiveView.Activate(GetDesktopWindow())
        Catch e As Exception
            sw.WriteLine("Error: document would not open. " & e.ToString)
            Exit Sub
        End Try

        Dim pMap As IMap
        Dim pLayer As ILayer
        'Dim pGeoFL As IGeoFeatureLayer
        Dim pEnumLayer As IEnumLayer
        Dim lMapCount As Integer

        lMapCount = pMapDocument.MapCount
        mxdProps.lMapCount = pMapDocument.MapCount
        Dim sDistUnits(pMapDocument.MapCount) As String
        ReDim mxdProps.sMapUnits(pMapDocument.MapCount)
        ReDim mxdProps.sSRef(pMapDocument.MapCount)

        Dim bMissing As Boolean
        Dim lRevision, lMajor, lMinor, lBuild As Integer
        pMapDocument.GetVersionInfo(bMissing, lMajor, lMinor, lRevision, lBuild)
        If Not bMissing Then sw.WriteLine("Document Version: " & lMajor & "." & lMinor & "." & lRevision & "." & lBuild)
        Select Case pMapDocument.DocumentVersion
            Case esriMapDocumentVersionInfo.esriMapDocumentVersionInfoSuccess
                sw.WriteLine("Document can be read at this version")
            Case esriMapDocumentVersionInfo.esriMapDocumentVersionInfoFail
                sw.WriteLine("Document cannot be read at this version")
            Case esriMapDocumentVersionInfo.esriMapDocumentVersionInfoUnknown
                sw.WriteLine("Document version not known")
            Case Else
                sw.WriteLine("Error: document version not known")
        End Select

        'file attributes
        sw.WriteLine("Last Modified: " & File.GetLastWriteTime(sMxdName))

        'doc paths
        If pMapDocument.UsesRelativePaths Then
            mxdProps.bRelPaths = True
            sw.WriteLine("Relative paths")
        Else
            mxdProps.bAbsPaths = True
            sw.WriteLine("Absolute paths")
        End If

        'layout
        Dim pPageLayout As IPageLayout
        If TypeOf pActiveView Is IPageLayout _
            Or TypeOf pActiveView Is IPageLayout2 _
            Or TypeOf pActiveView Is IPageLayout3 Then
            pPageLayout = pActiveView
            mxdProps.bLayoutView = True
            sw.WriteLine("Layout view")
        Else
            pPageLayout = pMapDocument.PageLayout
            mxdProps.bLayoutView = False
            sw.WriteLine("Data view")
        End If

        sw.WriteLine("Map Count: " & lMapCount)
        Dim pMapClipOptions As IMapClipOptions
        Dim pLayerSet As ISet
        Dim pDisplayTrans As IDisplayTransformation
        Dim pEnv As IEnvelope
        Dim pBarriers As IBarrierCollection = Nothing
        Dim pGeomColl As IGeometryCollection = Nothing
        Dim eWeight As esriBasicOverposterWeight
        Dim pMapFrame As IMapFrame
        Dim pSymBkg As ISymbolBackground
        Dim pGraphContainer As IGraphicsContainer
        Dim pElement As IElement
        Dim pMapAutoExtOpts As IMapAutoExtentOptions
        Dim dW, dH As Double
        Dim pMapOverposter As IMapOverposter
        Dim pOverposterOptions As IOverposterOptions = Nothing
        Dim pMaplexOverposterProperties As IMaplexOverposterProperties
        Dim pDictionary As IMaplexDictionary
        Dim pDictEntry As IMaplexDictionaryEntry
        Dim lNoEntries As Integer
        Dim pMapBookmarks As IMapBookmarks
        Dim pEnumSpatialBookmarks As IEnumSpatialBookmark
        Dim pAOIBookmark As IAOIBookmark
        Dim lLayerCount As Integer
        While lMapCount > 0
            If bAllLayers And Not mxdProps.bLayoutView Then
                pMap = pMapDocument.Map(CInt(pMapDocument.MapCount - lMapCount))
            Else
                pMap = pActiveView.FocusMap
            End If
            sw.WriteLine("Map #" & CShort(pMapDocument.MapCount - lMapCount + 1))
            'general map info
            sw.WriteLine(InsertTabs(1) & "Map Name: " & pMap.Name)
            If pMap.Description <> "" Then sw.WriteLine(InsertTabs(1) & "Description: " & pMap.Description)
            If pMap.AnnotationEngine Is Nothing Then
                'default to MLE if not found
                mxdProps.bMapIsMLE = True
            Else
                sw.WriteLine(InsertTabs(1) & "Label Engine: " & pMap.AnnotationEngine.Name)
                'detect SLE - don't try to read MLE props
                mxdProps.bMapIsMLE = False
                mxdProps.bMapIsSLE = False
                Dim SLEName As String = ""
                If StrComp(pMap.AnnotationEngine.Name, "ESRI Maplex Label Engine", CompareMethod.Text) = 0 Then
                    mxdProps.bMLE = True
                    mxdProps.bMapIsMLE = True
                End If
                'TODO check this against other versions
                'looks like ESRI was removed between 10.0 and 10.2
                If m_Version < 101 Then
                    SLEName = "ESRI Standard Label Engine"
                Else
                    SLEName = "Standard Label Engine"
                End If
                If StrComp(pMap.AnnotationEngine.Name, SLEName, CompareMethod.Text) = 0 Then
                    mxdProps.bSLE = True
                    mxdProps.bMapIsSLE = True
                End If
            End If

            If pMap.ReferenceScale.CompareTo(0) <> 0 Then
                sw.WriteLine(InsertTabs(1) & "Reference Scale: 1:" & pMap.ReferenceScale)
                mxdProps.bRefScale = True
            Else
                sw.WriteLine(InsertTabs(1) & "Reference Scale: None")
            End If
            If Not pMap.SpatialReference Is Nothing Then
                mxdProps.sSRef(pMapDocument.MapCount - lMapCount) = pMap.SpatialReference.Name()
                sw.WriteLine(InsertTabs(1) & "Spatial Reference: " & mxdProps.sSRef(pMapDocument.MapCount - lMapCount))
                If TypeOf pMap.SpatialReference Is IProjectedCoordinateSystem Then mxdProps.bProjected = True
                If TypeOf pMap.SpatialReference Is IGeographicCoordinateSystem Then mxdProps.bGeographic = True
            End If
            Try
                If pMap.MapScale > 0 Then
                    sw.WriteLine(InsertTabs(1) & "Map Scale: 1:" & pMap.MapScale)
                Else
                    sw.WriteLine(InsertTabs(1) & "Map Scale: None")
                End If
            Catch
                sw.WriteLine(InsertTabs(1) & "Map Scale: Could not determine")
            End Try
            sw.WriteLine(InsertTabs(1) & "Symbol Levels Used: " & pMap.UseSymbolLevels)

            If m_Version >= 94 Then
                pMapClipOptions = pMap
                If Not pMapClipOptions.ClipGeometry Is Nothing Then
                    mxdProps.bClipExtent = True
                    sw.WriteLine(InsertTabs(1) & "Clip shape: " & GetGeomType(pMapClipOptions.ClipGeometry.GeometryType))
                    sw.WriteLine(InsertTabs(2) & "Coords: " & pMapClipOptions.ClipGeometry.Envelope.XMin & "," & pMapClipOptions.ClipGeometry.Envelope.YMin & "-" & pMapClipOptions.ClipGeometry.Envelope.XMax & "," & pMapClipOptions.ClipGeometry.Envelope.YMax)
                    If Not pMapClipOptions.ClipBorder Is Nothing Then sw.WriteLine(InsertTabs(2) & "Border")
                    'Print # InsertTabs(2) & "Clip Data: " & pMapClipOptions.ClipData
                    pLayerSet = pMapClipOptions.ClipFilter
                    If Not pLayerSet Is Nothing Then
                        sw.WriteLine(InsertTabs(2) & "Excluded layers:")
                        pLayerSet.Reset()
                        pLayer = pLayerSet.Next
                        Do While Not pLayer Is Nothing
                            sw.WriteLine(InsertTabs(3) & pLayer.Name)
                            pLayer = pLayerSet.Next
                            mxdProps.bExcludeLayers = True
                        Loop
                    End If
                    sw.WriteLine(InsertTabs(2) & "Clip Grid and Graticules: " & pMapClipOptions.ClipGridAndGraticules)
                    sw.WriteLine(InsertTabs(2) & "Clip Type: " & GetClipType((pMapClipOptions.ClipType)))
                    If pMapClipOptions.ClipType = esriMapClipType.esriMapClipShape Then mxdProps.bClipToShape = True
                End If 'clipoptions
            End If '>=9.4
            If Not mxdProps.bClipExtent Then
                If Not pMap.ClipGeometry Is Nothing Then
                    mxdProps.bClipExtent = True
                    sw.WriteLine(InsertTabs(1) & "Clip shape: " & GetGeomType(pMap.ClipGeometry.GeometryType))
                    If pMap.ClipGeometry.GeometryType = esriGeometryType.esriGeometryBag Then mxdProps.bClipToShape = True
                    sw.WriteLine(InsertTabs(2) & "Coords: " & pMap.ClipGeometry.Envelope.XMin & "," & pMap.ClipGeometry.Envelope.YMin & "-" & pMap.ClipGeometry.Envelope.XMax & "," & pMap.ClipGeometry.Envelope.YMax)
                    If Not pMap.ClipBorder Is Nothing Then sw.WriteLine(InsertTabs(2) & "Border")
                End If 'clipgeometry
            End If 'bclipextent

            'map units
            mxdProps.sMapUnits(pMapDocument.MapCount - lMapCount) = GetUnits((pMap.MapUnits))
            sDistUnits(pMapDocument.MapCount - lMapCount) = GetUnits((pMap.DistanceUnits))
            sw.WriteLine(InsertTabs(1) & "Map Units: " & mxdProps.sMapUnits(pMapDocument.MapCount - lMapCount))
            sw.WriteLine(InsertTabs(1) & "Distance Units: " & sDistUnits(pMapDocument.MapCount - lMapCount))

            'barriers
            pDisplayTrans = pActiveView.ScreenDisplay.DisplayTransformation()
            pEnv = pDisplayTrans.Bounds
            pBarriers = pMap.Barriers(pEnv)
            sw.WriteLine(InsertTabs(1) & "Number of Barriers: " & pBarriers.Count)
            mxdProps.lBarriers = pBarriers.Count
            If bAllLayers Then
                For i = 0 To pBarriers.Count - 1
                    pBarriers.QueryItem(i, pGeomColl, eWeight)
                    sw.WriteLine(InsertTabs(2) & "#" & i + 1 & ": Weight = " & GetOverposterWeight(eWeight))
                    For j = 0 To pGeomColl.GeometryCount - 1
                        sw.WriteLine(InsertTabs(3) & "Item " & j & " type = " & GetGeomType(pGeomColl.Geometry(j).GeometryType))
                    Next
                    'Else 'not balllayers
                    '  sw.WriteLine(InsertTabs(3) & pGeomColl.GeometryCount & " items")
                Next
            End If
            'map anno
            sw.WriteLine(InsertTabs(1) & "Basic Graphics Layer:")
            GetLayerProps(pMap.BasicGraphicsLayer, 2)
            'selected features
            If pMap.SelectionCount Then sw.WriteLine(InsertTabs(1) & "Selected features: " & pMap.SelectionCount)

            'bounds etc.
            Dim pSymBorder As ISymbolBorder
            pGraphContainer = pMapDocument.PageLayout
            If m_Version < 94 Then
                pGraphContainer.Reset()
                pElement = pGraphContainer.Next
                While Not pElement Is Nothing
                    If TypeOf pElement Is IMapFrame Then
                        pMapFrame = pElement
                        'extent info up to 9.3
                        'TODO check
                        If StrComp(pMapFrame.Map.Name, pMap.Name, vbTextCompare) = 0 Then _
                          GetExtentInfo(pMapFrame.ExtentType, pMapFrame.MapBounds, pMapFrame.MapScale)
                        pSymBkg = pMapFrame.Background()
                        If bReadSymbols Then
                            sw.WriteLine(InsertTabs(1) & "Dataframe:")
                            If Not pMapFrame.Border Is Nothing Then
                                pSymBorder = pMapFrame.Border
                                sw.WriteLine(InsertTabs(2) & "Border Name: " & pSymBorder.Name)
                                GetSymbolProps(pSymBorder.LineSymbol, 2, False)
                            End If
                            If Not pSymBkg Is Nothing Then
                                sw.WriteLine(InsertTabs(2) & "Background Name: " & pSymBkg.Name)
                                GetSymbolProps(pSymBkg.FillSymbol, 2, False)
                            End If
                        Else
                            If Not pSymBkg Is Nothing Then
                                If Not pSymBkg.FillSymbol Is Nothing Then
                                    sw.WriteLine(InsertTabs(1) & "Background color (RGB): " & GetRGB(pSymBkg.FillSymbol.Color))
                                    sw.WriteLine(InsertTabs(1) & "Background color (CMYK): " & GetCMYK(pSymBkg.FillSymbol.Color))
                                End If
                            End If
                        End If 'read symbols
                    End If
                    pElement = pGraphContainer.Next
                End While
            Else
                'extent info post 9.3
                pMapAutoExtOpts = pMap
                GetExtentInfo(pMapAutoExtOpts.AutoExtentType, pMapAutoExtOpts.AutoExtentBounds, pMapAutoExtOpts.AutoExtentScale)
                pMapFrame = pGraphContainer.FindFrame(pActiveView.FocusMap)
                pSymBkg = pMapFrame.Background()
                If bReadSymbols Then
                    sw.WriteLine(InsertTabs(1) & "Dataframe:")
                    If Not pMapFrame.Border Is Nothing Then
                        pSymBorder = pMapFrame.Border
                        sw.WriteLine(InsertTabs(2) & "Border Name: " & pSymBorder.Name)
                        GetSymbolProps(pSymBorder.LineSymbol, 2, False)
                    End If
                    If Not pSymBkg Is Nothing Then
                        sw.WriteLine(InsertTabs(2) & "Background Name: " & pSymBkg.Name)
                        GetSymbolProps(pSymBkg.FillSymbol, 2, False)
                    End If
                Else
                    If Not pSymBkg Is Nothing Then
                        If Not pSymBkg.FillSymbol Is Nothing Then
                            sw.WriteLine(InsertTabs(1) & "Background color (RGB): " & GetRGB(pSymBkg.FillSymbol.Color))
                            sw.WriteLine(InsertTabs(1) & "Background color (CMYK): " & GetCMYK(pSymBkg.FillSymbol.Color))
                        End If
                    End If
                End If 'read symbols
            End If
            sw.WriteLine(InsertTabs(1) & "Map Bounds: " & pEnv.XMin & ", " & pEnv.YMin & ", " & pEnv.XMax & ", " & pEnv.YMax)
            pEnv = pDisplayTrans.VisibleBounds
            sw.WriteLine(InsertTabs(1) & "Visible Bounds: " & pEnv.XMin & ", " & pEnv.YMin & ", " & pEnv.XMax & ", " & pEnv.YMax)
            sw.WriteLine(InsertTabs(1) & "Display Transform Units: " & GetUnits((pDisplayTrans.Units)))

            'in page layout, bounds units are different to bookmark units.
            'try to convert (not working)
            '...
            '...
            'removed

            sw.WriteLine(InsertTabs(1) & "Rotation: " & pDisplayTrans.Rotation)
            If pDisplayTrans.Rotation.CompareTo(0) <> 0 Then mxdProps.bFrameRotation = True

            'get page size
            pMap.GetPageSize(dW, dH)
            If dW.CompareTo(0) <> 0 And dH.CompareTo(0) <> 0 Then sw.WriteLine(InsertTabs(1) & "Page Size: " & CDbl(dW) & " x " & CDbl(dH) & " inches")

            If bReadLabels Then
                'overposter options
                sw.WriteLine(InsertTabs(1) & "Overposter options:")
                pMapOverposter = pMap
                pOverposterOptions = pMapOverposter.OverposterProperties
                If pOverposterOptions.EnableDrawUnplaced Then
                    sw.WriteLine(InsertTabs(2) & "Draw unplaced")
                    mxdProps.bDrawUnplaced = True
                End If
                If pOverposterOptions.RotateLabelWithDataFrame Then
                    sw.WriteLine(InsertTabs(2) & "Rotate labels with dataframe")
                    mxdProps.bRotateWithDataFrame = True
                End If
                sw.WriteLine(InsertTabs(2) & "Inverted label tolerance: " & pOverposterOptions.InvertedLabelTolerance)
                mxdProps.sInvertedLabTol = mxdProps.sInvertedLabTol & " " & pOverposterOptions.InvertedLabelTolerance.ToString

                'Maplex overposter props
                If mxdProps.bMapIsMLE Then
                    sw.WriteLine(InsertTabs(1) & "Maplex overposter props:")
                    sTmp = ""
                    pMaplexOverposterProperties = pMapOverposter.OverposterProperties

                    If pMaplexOverposterProperties.EnableConnection Then
                        Select Case pMaplexOverposterProperties.ConnectionType
                            Case esriMaplexConnectionType.esriMaplexMinimizeLabels
                                sConnType = "Minimize"
                                If m_Version < 101 Then mxdProps.bMinimize = True
                            Case esriMaplexConnectionType.esriMaplexUnambiguous
                                sConnType = "Unambiguous"
                                If m_Version < 101 Then mxdProps.bUnambiguous = True
                            Case Else
                                sConnType = "Error: unknown type"
                        End Select
                        sw.WriteLine(InsertTabs(2) & "Line connection: " & sConnType & sTmp)
                    End If
                    If pMaplexOverposterProperties.AllowBorderOverlap Then
                        sw.WriteLine(InsertTabs(2) & "Allow border overlap")
                        mxdProps.bAllowOverlap = True
                    End If
                    If pMaplexOverposterProperties.LabelLargestPolygon Then
                        sw.WriteLine(InsertTabs(2) & "Label largest polygon" & sTmp)
                        If m_Version < 101 Then mxdProps.bLargestOnly = True
                    End If
                    If pMaplexOverposterProperties.PlacementQuality = _
                        esriMaplexPlacementQuality.esriMaplexPlacementQualityHigh Then _
                        sw.WriteLine(InsertTabs(2) & "Best placement")
                    If pMaplexOverposterProperties.PlacementQuality = _
                        esriMaplexPlacementQuality.esriMaplexPlacementQualityMedium Then _
                        sw.WriteLine(InsertTabs(2) & "Medium placement")
                    If pMaplexOverposterProperties.PlacementQuality = _
                        esriMaplexPlacementQuality.esriMaplexPlacementQualityLow Then
                        sw.WriteLine(InsertTabs(2) & "Fast placement")
                        mxdProps.bFast = True
                    End If
                    mxdProps.pDictionaries = pMaplexOverposterProperties.Dictionaries
                    If mxdProps.pDictionaries.DictionaryCount Then
                        sw.WriteLine(InsertTabs(2) & mxdProps.pDictionaries.DictionaryCount & " dictionaries found:")
                        For i = 0 To mxdProps.pDictionaries.DictionaryCount - 1
                            pDictionary = mxdProps.pDictionaries.GetDictionary(i)
                            sw.WriteLine(InsertTabs(3) & pDictionary.Name)
                            lNoEntries = pDictionary.EntryCount
                            For j = 0 To lNoEntries - 1
                                pDictEntry = pDictionary.GetEntry(j)
                                Select Case pDictEntry.Type
                                    Case esriMaplexAbbrevType.esriMaplexAbbrevTypeEnding
                                        mxdProps.bDictionaryEnding = True
                                    Case esriMaplexAbbrevType.esriMaplexAbbrevTypeKeyword
                                        mxdProps.bDictionaryKeyword = True
                                    Case esriMaplexAbbrevType.esriMaplexAbbrevTypeTranslation
                                        mxdProps.bDictionaryTranslation = True
                                    Case Else
                                        sw.WriteLine(InsertTabs(3) & "Error: Dictionary entry type not known")
                                End Select
                            Next
                        Next
                    End If 'dictionaries
                End If 'mle

                'bookmarks
                pMapBookmarks = pMap
                pEnumSpatialBookmarks = pMapBookmarks.Bookmarks
                pEnumSpatialBookmarks.Reset()
                pAOIBookmark = pEnumSpatialBookmarks.Next

                If Not pAOIBookmark Is Nothing Then sw.WriteLine(InsertTabs(1) & "Bookmarks:")
                Do Until (pAOIBookmark Is Nothing)
                    Try
                        pEnv = pAOIBookmark.Location
                        sw.WriteLine(InsertTabs(2) & pAOIBookmark.Name() & " (" & pEnv.XMin & ", " & pEnv.YMin & ", " & pEnv.XMax & ", " & pEnv.YMax & ")")
                    Catch
                        sw.WriteLine(InsertTabs(2) & pAOIBookmark.Name() & " (Error: bounds invalid)")
                    End Try
                    'pAOIBookmark.ZoomTo pActiveView.FocusMap
                    'pActiveView.Refresh
                    'Set pEnv = pActiveView.extent
                    'Print # inserttabs(1) & "Activeview:  (" & pEnv.XMin & ", " & pEnv.YMin & ", " & pEnv.XMax & ", " & pEnv.YMax & ")"
                    pAOIBookmark = pEnumSpatialBookmarks.Next
                Loop
            End If 'read symbols

            'layer info
            sw.WriteLine(vbCrLf & InsertTabs(1) & "Number of Layers: " & pMap.LayerCount)
            If pMap.LayerCount > 0 Then
                sw.WriteLine(InsertTabs(1) & "Layer properties:")

                'pUID = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" 'IGeoFeatureLayer
                'pUID = "{EDAD6644-1810-11D1-86AE-0000F8751720}" 'IGroupLayer
                'pUID = "{5CEAE408-4C0A-437F-9DB3-054D83919850}" 'IFDOGraphicsLayer = Anno
                'pUID = "{6CA416B1-E160-11D2-9F4E-00C04F6BC78E}" 'IDataLayer = All except group
                'pUID = "{40A9E885-5533-11d0-98BE-00805F7CED21}" 'IFeatureLayer = FDO Graphics + GeoFeature
                pEnumLayer = pMap.Layers(Nothing, False) 'Nothing for all layers, pUID for specific layer
                pEnumLayer.Reset()
                pLayer = pEnumLayer.Next

                'look at each layer in the map
                Do While Not pLayer Is Nothing
                    lLayerCount = lLayerCount + 1
                    sw.WriteLine(InsertTabs(1) & pMap.Name & " layer " & lLayerCount & "/" & pMap.LayerCount)
                    GetLayerProps(pLayer, 2, pMap.UseSymbolLevels)
                    sw.WriteLine("")
                    pLayer = pEnumLayer.Next
                Loop  'feature layer
            Else
                sw.WriteLine("")
            End If

            'list tables
            Dim pTableColl As IStandaloneTableCollection = CType(pMap, IStandaloneTableCollection)
            Dim pTable As ITable
            Dim pDataset As IDataset
            For i = 0 To pTableColl.StandaloneTableCount - 1
                pTable = pTableColl.StandaloneTable(i)
                pDataset = CType(pTable, IDataset)
                sw.WriteLine(InsertTabs(1) & "Table " & i + 1 & "/" & pTableColl.StandaloneTableCount)
                If pDataset Is Nothing Then
                    sw.WriteLine(InsertTabs(2) & "WARNING: Could not read table details")
                    sw.WriteLine("")
                    Continue For
                End If
                sw.WriteLine(InsertTabs(2) & "Name: " & pDataset.Name)
                Try
                    sw.WriteLine(InsertTabs(2) & "Data source: " & pDataset.Workspace.PathName)
                    mxdProps.sDataSources(mxdProps.lDataSources) = pDataset.Workspace.PathName
                    AddIfUnique(mxdProps.lDataSources, mxdProps.sDataSources, ARRAY_SIZE)
                Catch ex As Exception
                    sw.WriteLine(InsertTabs(2) & "WARNING: Could not read data source name")
                End Try
                sw.WriteLine("")
            Next

            sw.WriteLine(String.Format("End Map #{0}", pMapDocument.MapCount - lMapCount + 1))
            If mxdProps.bLayoutView Or bAllLayers Then
                'layout view - next frame
                If lMapCount > 1 Then
                    pPageLayout.FocusNextMapFrame()
                End If
                lMapCount = lMapCount - 1
                lLayerCount = 0
            Else
                'data view - do not try to look at next map frame, just exit
                lMapCount = 0
            End If
        End While

        If bReadSymbols Then
            WriteSymbolSummary()
            If bExcel And File.Exists(sSummaryXls) Then WriteXLS(sMxdName)
        End If
        If bReadLabels Then
            WriteLabelSummary()
            If mxdProps.bMLE And bExcel And File.Exists(sSummaryXls) Then WriteXLS(sMxdName)
        End If
        WriteGeneralSummary()
        sw.WriteLine(vbCrLf & "OK")

    End Sub

    'Open archive file and read all the XML files inside
    Private Sub ReadXMLArchive(ByVal sFile As String, ByVal bExcel As Boolean)

        'create new set of props and temp directory
        mxdProps = New clsMxdProps
        Dim sTempPath As String = IO.Path.Combine(IO.Path.GetTempPath(), "ReadMxdXMLArchives")
        sTempPath = IO.Path.Combine(sTempPath, IO.Path.GetFileNameWithoutExtension(sFile))
        Dim lTabLevel As Long = 0

        If bExcel Then
            ' TODO excel export
        End If

        If bAllLayers Then sw.WriteLine("Showing all layers")
        'file attributes
        sw.WriteLine("Last Modified: " & File.GetLastWriteTime(sFile))
        'make sure temp dir is empty
        DeleteDirectory(sTempPath)
        Try
            Using zip As ZipFile = New ZipFile(sFile)
                zip.ExtractAll(sTempPath)
            End Using

            For Each xmlfile As String In Directory.GetFiles(sTempPath, "*.xml", SearchOption.AllDirectories)
                sw.WriteLine("{0}{1}:", InsertTabs(lTabLevel), xmlfile.Replace(sTempPath, ""))
                sw.Flush()
                Dim xtr As Xml.XmlTextReader = New Xml.XmlTextReader(xmlfile)
                Do While xtr.Read()
                    Select Case xtr.NodeType
                        Case Xml.XmlNodeType.Element
                            lTabLevel = lTabLevel + 1
                            If xtr.Name().Length > 0 Then sw.WriteLine("{0}{1}:", InsertTabs(lTabLevel), xtr.Name())
                        Case Xml.XmlNodeType.Text
                            If xtr.Value().Length > 0 Then sw.WriteLine("{0}{1}", InsertTabs(lTabLevel + 1), xtr.Value())
                        Case Xml.XmlNodeType.EndElement
                            lTabLevel = lTabLevel - 1
                    End Select
                Loop
                xtr.Close()
                sw.Flush()
            Next
        Catch e As Exception
            sw.WriteLine("Error reading archive." & e.ToString)
        Finally
            DeleteDirectory(sTempPath)
            sw.Flush()
        End Try

    End Sub

    'Sum up all the properties used in this map document
    Private Sub WriteLabelSummary()

        Dim sTmp As String
        Dim i As Integer
        sw.WriteLine(vbCrLf & "Label Properties Summary:" & vbCrLf & "Polygons:")
        If mxdProps.bPolyHorz Then sw.WriteLine(InsertTabs(1) & "Polygon horizontal")
        If mxdProps.bPolyStr Then sw.WriteLine(InsertTabs(1) & "Polygon straight")
        If mxdProps.bPolyCurv Then sw.WriteLine(InsertTabs(1) & "Polygon curved")
        If mxdProps.bPolyOffHorz Then sw.WriteLine(InsertTabs(1) & "Polygon offset horizontal")
        If mxdProps.bPolyOffCurv Then sw.WriteLine(InsertTabs(1) & "Polygon offset curved")
        If mxdProps.bPolyRegular Then sw.WriteLine(InsertTabs(1) & "Regular polygon")
        If mxdProps.bPolyParcel Then sw.WriteLine(InsertTabs(1) & "Land parcel")
        If mxdProps.bPolyRiver Then sw.WriteLine(InsertTabs(1) & "River polygon")
        If mxdProps.bPolyBdy Then sw.WriteLine(InsertTabs(1) & "Boundary")
        If mxdProps.bPolyBdySingleSided Then sw.WriteLine(InsertTabs(2) & "Single sided")
        If mxdProps.bPolyBdyAllowHoles Then sw.WriteLine(InsertTabs(2) & "Allow holes")
        If mxdProps.bPolyBdyOnLine Then sw.WriteLine(InsertTabs(2) & "On line")
        If mxdProps.bPolyTryHorz Then sw.WriteLine(InsertTabs(1) & "Try horizontal first")
        If mxdProps.bPolyMayPlaceOutside Then sw.WriteLine(InsertTabs(1) & "May place outside")
        If mxdProps.bPolyPlaceOnlyInside Then sw.WriteLine(InsertTabs(1) & "Only place inside")
        If mxdProps.bPolyOffDist Then sw.WriteLine(InsertTabs(1) & "Polygon offset distance")
        If mxdProps.bPolyMaxOffset Then sw.WriteLine(InsertTabs(1) & "Polygon maximum offset")
        If mxdProps.bPolyFtrGeom Then sw.WriteLine(InsertTabs(1) & "Polygon feature geometry")
        If mxdProps.bPolyGACurv Then sw.WriteLine(InsertTabs(1) & "Graticule alignment curved")
        If mxdProps.bPolyGACurvNoFlip Then sw.WriteLine(InsertTabs(1) & "Graticule alignment curved (no flip)")
        If mxdProps.bPolyGAStr Then sw.WriteLine(InsertTabs(1) & "Graticule alignment straight")
        If mxdProps.bPolyGAStrNoFlip Then sw.WriteLine(InsertTabs(1) & "Graticule alignment straight (no flip)")
        If mxdProps.bPolyIntZones Then sw.WriteLine(InsertTabs(1) & "Internal zones")
        If mxdProps.bPolyExtZones Then sw.WriteLine(InsertTabs(1) & "External zones")
        If mxdProps.bPolyAnchor Then sw.WriteLine(InsertTabs(1) & "Anchor points")
        If mxdProps.bPolyRepeat Then sw.WriteLine(InsertTabs(1) & "Polygon repeat")
        If mxdProps.bPolySpread Then sw.WriteLine(InsertTabs(1) & "Spread chars")
        If mxdProps.bPolyAllowHoles Then sw.WriteLine(InsertTabs(1) & "Allow holes")
        If mxdProps.bMultipatch Then sw.WriteLine(InsertTabs(1) & "Multipatch")
        sw.WriteLine("Lines:")
        If mxdProps.bLineCenHor Then sw.WriteLine(InsertTabs(1) & "Line centered horizontal")
        If mxdProps.bLineCenStr Then sw.WriteLine(InsertTabs(1) & "Line centered straight")
        If mxdProps.bLineCenCur Then sw.WriteLine(InsertTabs(1) & "Line centered curved")
        If mxdProps.bLineCenPer Then sw.WriteLine(InsertTabs(1) & "Line centered perpendicular")
        If mxdProps.bLineOffHor Then sw.WriteLine(InsertTabs(1) & "Line offset horizontal")
        If mxdProps.bLineOffStr Then sw.WriteLine(InsertTabs(1) & "Line offset straight")
        If mxdProps.bLineOffCur Then sw.WriteLine(InsertTabs(1) & "Line offset curved")
        If mxdProps.bLineOffPer Then sw.WriteLine(InsertTabs(1) & "Line offset perpendicular")
        If mxdProps.bLineHor Then sw.WriteLine(InsertTabs(1) & "Line horizontal")
        If mxdProps.bLineParallel Then sw.WriteLine(InsertTabs(1) & "Line parallel")
        If mxdProps.bLineCrv Then sw.WriteLine(InsertTabs(1) & "Line curved")
        If mxdProps.bLinePerp Then sw.WriteLine(InsertTabs(1) & "Line perpendicular")
        If mxdProps.bLineRegular Then sw.WriteLine(InsertTabs(1) & "Regular line")
        If mxdProps.bLineStreet Then
            sw.WriteLine(InsertTabs(1) & "Street")
            If mxdProps.bStreetHorz Then sw.WriteLine(InsertTabs(2) & "Horizontal")
            If mxdProps.bStreetReduce Then sw.WriteLine(InsertTabs(2) & "Reduce leading")
            If mxdProps.bStreetPrimary Then sw.WriteLine(InsertTabs(2) & "Primary name under")
            If mxdProps.bStreetSpread Then sw.WriteLine(InsertTabs(2) & "Spread words")
        End If
        If mxdProps.bLineStreetAdd Then sw.WriteLine(InsertTabs(1) & "Street address")
        If mxdProps.bLineContour Then
            sw.WriteLine(InsertTabs(1) & "Contour")
            If mxdProps.bContourPage Then sw.WriteLine(InsertTabs(2) & "Page alignment")
            If mxdProps.bContourUphill Then sw.WriteLine(InsertTabs(2) & "Uphill alignment")
            If mxdProps.bContourNoLadder Then sw.WriteLine(InsertTabs(2) & "No ladders")
            If mxdProps.bContourLadder Then sw.WriteLine(InsertTabs(2) & "Ladders")
        End If
        If mxdProps.bLineRiver Then sw.WriteLine(InsertTabs(1) & "River line")
        If mxdProps.bLineSecOff Then sw.WriteLine(InsertTabs(1) & "Line secondary offset")
        If mxdProps.bLineOffDist Or mxdProps.bLineOffset Then sw.WriteLine(InsertTabs(1) & "Line offset distance")
        'If mxdProps.bLinePrefOffset Then sw.WriteLine(InsertTabs(1) & "Line preferred offset")
        If mxdProps.bConstrainAbove Then sw.WriteLine(InsertTabs(1) & "Line constrain above")
        If mxdProps.bConstrainBelow Then sw.WriteLine(InsertTabs(1) & "Line constrain below")
        If mxdProps.bConstrainLeft Then sw.WriteLine(InsertTabs(1) & "Line constrain left")
        If mxdProps.bConstrainRight Then sw.WriteLine(InsertTabs(1) & "Line constrain right")
        If mxdProps.bNoConstraint Then sw.WriteLine(InsertTabs(1) & "Line no constraint")
        If mxdProps.bLineFtrGeom Then sw.WriteLine(InsertTabs(1) & "Line feature geometry")
        If mxdProps.bLineBestAlong Then sw.WriteLine(InsertTabs(1) & "Best along line ")
        If mxdProps.bLineBeforeStart Then sw.WriteLine(InsertTabs(1) & "Line before start")
        If mxdProps.bLineAfterEnd Then sw.WriteLine(InsertTabs(1) & "Line after end")
        If mxdProps.bLineFromStart Then sw.WriteLine(InsertTabs(1) & "Line along from start")
        If mxdProps.bLineFromEnd Then sw.WriteLine(InsertTabs(1) & "Line along from end")
        If mxdProps.bStraddlacking Then sw.WriteLine(InsertTabs(1) & "Allow straddle stacking")
        If mxdProps.bLineGACurv Then sw.WriteLine(InsertTabs(1) & "Graticule alignment curved")
        If mxdProps.bLineGACurvNoFlip Then sw.WriteLine(InsertTabs(1) & "Graticule alignment curved (no flip)")
        If mxdProps.bLineGAStr Then sw.WriteLine(InsertTabs(1) & "Graticule alignment straight")
        If mxdProps.bLineGAStrNoFlip Then sw.WriteLine(InsertTabs(1) & "Graticule alignment straight (no flip)")
        If mxdProps.bLineDirection Then sw.WriteLine(InsertTabs(1) & "Align to direction")
        If mxdProps.bLineRepeat Then sw.WriteLine(InsertTabs(1) & "Line repeat")
        If mxdProps.bLabelNearBorder Then sw.WriteLine(InsertTabs(1) & "Prefer label near border")
        If mxdProps.bLabelNearJunction Then sw.WriteLine(InsertTabs(1) & "Prefer label near junctions")
        If mxdProps.bLineSpread Then sw.WriteLine(InsertTabs(1) & "Line spread chars")
        If mxdProps.bMultiOptionFeature Then sw.WriteLine(InsertTabs(1) & "One label per feature")
        If mxdProps.bMultiOptionPart Then sw.WriteLine(InsertTabs(1) & "One label per part")
        If mxdProps.bMultiOptionSegment Then sw.WriteLine(InsertTabs(1) & "One label per segment")
        sw.WriteLine("Points:")
        If mxdProps.bPointFixed Then sw.WriteLine(InsertTabs(1) & "Point fixed")
        If mxdProps.bMayShift Then sw.WriteLine(InsertTabs(1) & "May shift")
        If mxdProps.bPointBest Then sw.WriteLine(InsertTabs(1) & "Point best")
        If mxdProps.bPointAround Then sw.WriteLine(InsertTabs(1) & "Place around point")
        If mxdProps.bPointOnTop Then sw.WriteLine(InsertTabs(1) & "Place on top of point")
        If mxdProps.bAlteredZones Then
            sw.WriteLine(InsertTabs(1) & "Point zones (altered)")
        ElseIf mxdProps.bPointZones Then
            sw.WriteLine(InsertTabs(1) & "Point zones (default)")
        End If
        If mxdProps.bPointOffDist Then sw.WriteLine(InsertTabs(1) & "Point offset distance")
        If mxdProps.bPointMaxOffset Then sw.WriteLine(InsertTabs(1) & "Point maximum offset")
        If mxdProps.bPointFtrGeom Then sw.WriteLine(InsertTabs(1) & "Point feature geometry")
        If mxdProps.bSymbolOutline Then sw.WriteLine(InsertTabs(1) & "Point symbol outline")
        If mxdProps.bPointGACurv Then sw.WriteLine(InsertTabs(1) & "Graticule alignment curved")
        If mxdProps.bPointGACurvNoFlip Then sw.WriteLine(InsertTabs(1) & "Graticule alignment curved (no flip)")
        If mxdProps.bPointGAStr Then sw.WriteLine(InsertTabs(1) & "Graticule alignment straight")
        If mxdProps.bPointGAStrNoFlip Then sw.WriteLine(InsertTabs(1) & "Graticule alignment straight (no flip)")
        If mxdProps.bPointRotation Then sw.WriteLine(InsertTabs(1) & "Point rotation")
        If mxdProps.bPointRotAngle Then sw.WriteLine(InsertTabs(2) & "Additional angle")
        If mxdProps.bPointSpecAngle Then sw.WriteLine(InsertTabs(2) & "Specified angle")
        If mxdProps.bPointRotFlip Then sw.WriteLine(InsertTabs(2) & "May flip")
        If mxdProps.bMultipoint Then sw.WriteLine(InsertTabs(1) & "Multipoint")
        If mxdProps.bWhitespace Then sw.WriteLine(InsertTabs(1) & "Remove whitespace off")
        If mxdProps.bLinebreaks Then sw.WriteLine(InsertTabs(1) & "Remove line breaks")
        If mxdProps.bSLE Then sw.WriteLine("Num lab options:")
        If mxdProps.bNumLabNoRestrict Then sw.WriteLine(InsertTabs(1) & "No label restrictions")
        If mxdProps.bNumLabperName Then sw.WriteLine(InsertTabs(1) & "One label per feature")
        If mxdProps.bNumLabperPart Then sw.WriteLine(InsertTabs(1) & "One label per feature part")
        If mxdProps.bRemoveDupSLE Then sw.WriteLine(InsertTabs(1) & "Remove duplicate labels")
        If mxdProps.bMLE Then sw.WriteLine("Strategies:")
        If mxdProps.bStack Then
            sw.WriteLine(InsertTabs(1) & "Stack")
            If mxdProps.bStackC Then sw.WriteLine(InsertTabs(1) & "Stack center")
            If mxdProps.bStackR Then sw.WriteLine(InsertTabs(1) & "Stack right")
            If mxdProps.bStackL Then sw.WriteLine(InsertTabs(1) & "Stack left")
            If mxdProps.bStackLorR Then sw.WriteLine(InsertTabs(1) & "Stack left or right")
            sTmp = ""
            For i = 0 To mxdProps.lSeparators - 1
                sTmp = sTmp & "'" & mxdProps.sSeparators(i) & "' "
            Next
            If Len(sTmp) > 12 Or InStr(1, sTmp, "' '") = 0 Or InStr(1, sTmp, "','") = 0 Or _
                InStr(1, sTmp, "'-'") = 0 Then sw.WriteLine(InsertTabs(1) & "Stacking separators: " & sTmp)
            If mxdProps.bMaxLines Then sw.WriteLine(InsertTabs(1) & "Maximum lines <> 3")
            If mxdProps.bMinChars Then sw.WriteLine(InsertTabs(1) & "Minimim chars <> 3")
            If mxdProps.bMaxChars Then sw.WriteLine(InsertTabs(1) & "Maximum chars <> 24")
        End If
        If mxdProps.bOverrun Then sw.WriteLine(InsertTabs(1) & "Overrun")
        If mxdProps.bAsymmetric Then sw.WriteLine(InsertTabs(1) & "Asymmetric")
        If mxdProps.bFontReduction Then sw.WriteLine(InsertTabs(1) & "Font Reduction")
        If mxdProps.bCompression Then sw.WriteLine(InsertTabs(1) & "Compression")
        If mxdProps.bAbbreviation Then
            sw.WriteLine(InsertTabs(1) & "Abbreviation")
            If mxdProps.bDictionaryTranslation Then sw.WriteLine(InsertTabs(2) & "Translation")
            If mxdProps.bDictionaryKeyword Then sw.WriteLine(InsertTabs(2) & "Keyword")
            If mxdProps.bDictionaryEnding Then sw.WriteLine(InsertTabs(2) & "Ending")
        End If
        If mxdProps.bTruncation Then sw.WriteLine(InsertTabs(1) & "Truncation")
        If mxdProps.bTruncationLength Then sw.WriteLine(InsertTabs(2) & "Minimum length")
        If mxdProps.bTruncationMarker Then
            sTmp = ""
            For i = 0 To mxdProps.lTruncMarker - 1
                sTmp = sTmp & "'" & mxdProps.sTruncMarker(i) & "' "
            Next
            sw.WriteLine(InsertTabs(2) & "Truncation markers: " & sTmp)
        End If
        If mxdProps.bTruncationChars Then
            sTmp = ""
            For i = 0 To mxdProps.lTruncChars - 1
                sTmp = sTmp & "'" & mxdProps.sTruncChars(i) & "' "
            Next
            sw.WriteLine(InsertTabs(2) & "Preferred chars: " & sTmp)
        End If
        If mxdProps.bMinSize Then sw.WriteLine(InsertTabs(1) & "Min Size")
        If mxdProps.bKeyNumbering Then sw.WriteLine(InsertTabs(1) & "Key numbering")
        If mxdProps.bStrategyPriority Then sw.WriteLine(InsertTabs(1) & "Strategy priority order")
        sw.WriteLine("Conflicts:")
        If mxdProps.bWeights Then sw.WriteLine(InsertTabs(1) & "Weights")
        If mxdProps.bBackground Then sw.WriteLine(InsertTabs(1) & "Background")
        If mxdProps.bRemoveDup Then sw.WriteLine(InsertTabs(1) & "Remove duplicates")
        If mxdProps.bLabelBuffer Then sw.WriteLine(InsertTabs(1) & "Label buffer")
        If mxdProps.bHardConstraint Then sw.WriteLine(InsertTabs(2) & "Hard constraint")
        If mxdProps.bNeverRemove Then sw.WriteLine(InsertTabs(1) & "Never remove")
        If mxdProps.bOverlappingLabels Then sw.WriteLine(InsertTabs(1) & "Allow overlapping labels")
        sw.WriteLine("Text Symbol:")
        If mxdProps.bXYOffset Then sw.WriteLine(InsertTabs(1) & "XY offset")
        If mxdProps.bRighttoLeft Then sw.WriteLine(InsertTabs(1) & "Right to Left")
        If mxdProps.bTextPosition Then sw.WriteLine(InsertTabs(1) & "Text position")
        If mxdProps.bTextCase Then sw.WriteLine(InsertTabs(1) & "Text case")
        If mxdProps.bCharSpacing Then sw.WriteLine(InsertTabs(1) & "Char spacing")
        If mxdProps.bLeading Then sw.WriteLine(InsertTabs(1) & "Leading")
        If mxdProps.bCharWidth Then sw.WriteLine(InsertTabs(1) & "Char width")
        If mxdProps.bWordSpacing Then sw.WriteLine(InsertTabs(1) & "Word spacing")
        If mxdProps.bKerningOff Then sw.WriteLine(InsertTabs(1) & "Kerning off")
        If mxdProps.bFillSymbol Then sw.WriteLine(InsertTabs(1) & "Fill Symbol")
        If mxdProps.bTextBackground Then sw.WriteLine(InsertTabs(1) & "Text background")
        If mxdProps.bBalloonCallout Then sw.WriteLine(InsertTabs(2) & "Balloon callout")
        If mxdProps.bLineCallout Then sw.WriteLine(InsertTabs(2) & "Line callout")
        If mxdProps.bMarkerTextBkg Then sw.WriteLine(InsertTabs(2) & "Marker text background")
        If mxdProps.bScaletoFit Then sw.WriteLine(InsertTabs(3) & "Scale to fit")
        If mxdProps.bSimpleLineCallout Then sw.WriteLine(InsertTabs(2) & "Simple line callout")
        If mxdProps.bShadow Then sw.WriteLine(InsertTabs(1) & "Shadow")
        If mxdProps.bHalo Then sw.WriteLine(InsertTabs(1) & "Halo")
        If mxdProps.bCJK Then sw.WriteLine(InsertTabs(1) & "CJK")
        sw.WriteLine("Misc:")
        If mxdProps.bLayerDefQuery Then sw.WriteLine(InsertTabs(1) & "Layer Definition Query")
        If mxdProps.bSQL Then sw.WriteLine(InsertTabs(1) & "SQL Query")
        If mxdProps.bLabelExpression Then sw.WriteLine(InsertTabs(1) & "Label Expression")
        If mxdProps.bBaseTag Then sw.WriteLine(InsertTabs(2) & "Base Tag")
        If mxdProps.bTags Then sw.WriteLine(InsertTabs(2) & "Tags")
        If mxdProps.bHTMLEnt Then sw.WriteLine(InsertTabs(2) & "HTML entities")
        If mxdProps.bCodedValueDomain Then sw.WriteLine(InsertTabs(1) & "Coded value domain")
        If mxdProps.bScaleRanges Then sw.WriteLine(InsertTabs(1) & "Scale ranges")
        If mxdProps.bDrawUnplaced Then sw.WriteLine(InsertTabs(1) & "Draw unplaced")
        If mxdProps.bRotateWithDataFrame Then sw.WriteLine(InsertTabs(1) & "Rotate labels with dataframe")
        sw.WriteLine(InsertTabs(1) & "Inverted label tolerance:" & mxdProps.sInvertedLabTol)
        If mxdProps.bUnambiguous Then sw.WriteLine(InsertTabs(1) & "Unambiguous")
        If mxdProps.bMinimize Then sw.WriteLine(InsertTabs(1) & "Minimize")
        If mxdProps.bAllowOverlap Then sw.WriteLine(InsertTabs(1) & "Allow border overlap")
        If mxdProps.bLargestOnly Then sw.WriteLine(InsertTabs(1) & "Label largest only")
        If mxdProps.bFrameRotation Then sw.WriteLine(InsertTabs(1) & "Dataframe rotation")
        If mxdProps.lLabelClassCount > 1 Then
            If mxdProps.bLabelPriority Then sw.WriteLine(InsertTabs(1) & "Label priority ranking")
            If mxdProps.bUninitPriority Then sw.WriteLine(InsertTabs(1) & "UNINITIALISED LABEL PRIORITIES!")
        End If
        sw.WriteLine(InsertTabs(1) & "No of anno layers: " & mxdProps.lAnnoLayers)
        sw.WriteLine(InsertTabs(1) & "No of barriers: " & mxdProps.lBarriers)
        If mxdProps.bFast Then sw.WriteLine(InsertTabs(1) & "Fast placement")
        sw.WriteLine("No of label classes: " & mxdProps.lLabelClassCount)
        sw.WriteLine("No of label classes in anno layers: " & mxdProps.lAnnoLabelClassCount)
        If mxdProps.bMLE Then sw.WriteLine("MLE")
        If mxdProps.bSLE Then sw.WriteLine("SLE")

    End Sub

    Private Sub WriteGeneralSummary()

        Dim i As Integer
        sw.WriteLine(vbCrLf & "General Properties Summary:")
        If mxdProps.bLayoutView Then sw.WriteLine("Layout view") Else sw.WriteLine("Data view")
        If mxdProps.bLayoutView And mxdProps.lMapCount > 1 Then
            For i = 0 To mxdProps.lMapCount - 1
                sw.WriteLine("Map #" & i + 1 & ":")
                sw.WriteLine(InsertTabs(1) & "Spatial Reference: " & mxdProps.sSRef(i))
                If mxdProps.bGeographic Then sw.WriteLine(InsertTabs(1) & "Geographic coordinate system")
                If mxdProps.bProjected Then sw.WriteLine(InsertTabs(1) & "Projected coordinate system")
                sw.WriteLine(InsertTabs(1) & "Map Units: " & mxdProps.sMapUnits(i))
            Next
        ElseIf Not bLyrFile Then
            If mxdProps.lMapCount > 1 Then sw.WriteLine("Multiple dataframes")
            sw.WriteLine("Spatial Reference: " & mxdProps.sSRef(0))
            If mxdProps.bGeographic Then sw.WriteLine("Geographic coordinate system")
            If mxdProps.bProjected Then sw.WriteLine("Projected coordinate system")
            sw.WriteLine("Map Units: " & mxdProps.sMapUnits(0))
        End If
        If mxdProps.bRefScale Then sw.WriteLine("Reference scale")
        If mxdProps.bFixedExtent Then sw.WriteLine("Map Frame Fixed Extent")
        If mxdProps.bFixedScale Then sw.WriteLine("Map Frame Fixed Scale")
        If mxdProps.bAutoExtent Then sw.WriteLine("Map Frame Automatic Extent")
        If mxdProps.bClipExtent Then
            sw.WriteLine("Clip extent")
            If mxdProps.bClipToShape Then sw.WriteLine(InsertTabs(1) & "Clip to shape")
            If mxdProps.bExcludeLayers Then sw.WriteLine(InsertTabs(1) & "Exclude layers")
        End If
        If mxdProps.lDataSources > 1 Then
            sw.WriteLine("Data sources:")
            For i = 0 To mxdProps.lDataSources - 1
                sw.WriteLine(InsertTabs(1) & mxdProps.sDataSources(i))
            Next
        Else
            sw.WriteLine("Data source: " & mxdProps.sDataSources(0))
        End If
        If mxdProps.bSHP Then sw.WriteLine("Shapefile")
        If mxdProps.bFGDB Then sw.WriteLine("File Geodatabase")
        If mxdProps.bPGDB Then sw.WriteLine("Personal Geodatabase")
        If mxdProps.bSDE Then sw.WriteLine("SDE")
        If mxdProps.bCoverage Then sw.WriteLine("Coverage")
        If mxdProps.bRelPaths Then sw.WriteLine("Relative paths")
        If mxdProps.bAbsPaths Then sw.WriteLine("Absolute paths")
        If mxdProps.bQualifiedNames Then sw.WriteLine("Qualified names")

    End Sub

    'Sum up all the symbols used in this map document
    Private Sub WriteSymbolSummary()

        'Dim sTmp As String
        'Dim i As Integer
        sw.WriteLine(vbCrLf & "Symbol Properties Summary:")
        If mxdProps.bColorRamp Then sw.WriteLine("Color ramp")
        If mxdProps.bRasterClassify Then sw.WriteLine("Raster classify color ramp renderer")
        If mxdProps.bRasterStretch Then sw.WriteLine("Raster stretch color ramp renderer")
        If mxdProps.bRasterDiscrete Then sw.WriteLine("Raster discrete color renderer")
        If mxdProps.bRasterUnique Then sw.WriteLine("Raster unique value renderer")
        If mxdProps.bRasterRGB Then sw.WriteLine("Raster RGB renderer")
        If mxdProps.bBarChart Then sw.WriteLine("Bar chart")
        If mxdProps.bStackedChart Then sw.WriteLine("Stacked chart")
        If mxdProps.bPieChart Then sw.WriteLine("Pie chart")
        If mxdProps.b3DChart Then sw.WriteLine("3D chart")
        If mxdProps.bChartOverlap Then sw.WriteLine("Chart overlap")
        If mxdProps.bFixedSize Then sw.WriteLine("Fixed size")
        If mxdProps.bChartLeaders Then sw.WriteLine("Leaders")
        If mxdProps.bSimpleFill Then sw.WriteLine("Simple fill")
        If mxdProps.bLineFill Then sw.WriteLine("Line fill")
        If mxdProps.bMarkerFill Then sw.WriteLine("Marker fill")
        If mxdProps.bGradientFill Then sw.WriteLine("Gradient fill")
        If mxdProps.bPictureFill Then sw.WriteLine("Picture fill")
        If mxdProps.bBarOrient Then sw.WriteLine("Bar orientation")
        If mxdProps.bColumnOrient Then sw.WriteLine("Column orientation")
        If mxdProps.bArithOrient Then sw.WriteLine("Arithmetic orientation")
        If mxdProps.bGeogOrient Then sw.WriteLine("Geographic orientation")

	End Sub

    '***************** add summary to spreadsheet *****************
    Private Sub WriteXLS(ByVal sMxdName As String)
        Dim i, lCol As Integer
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlWbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim xlSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim bPolygons, bGraticule, bLines, bPoints, bStrategies, bDone As Boolean
        Dim lRow As Long = 3
        Dim lMaxRows As Long
        Dim lMaxCols As Long
        Dim lBlankRows As Long
        Dim sCellVal As String
        xlApp.DisplayAlerts = False

        Try
            xlWbook = xlApp.Workbooks.Open(sSummaryXls)
            xlSheet1 = xlWbook.Sheets.Item(1)
            lMaxCols = xlSheet1.Columns.CountLarge
            lMaxRows = xlSheet1.Rows.CountLarge

            lCol = xlSheet1.Range(xlSheet1.Cells(1, lMaxCols - 1), xlSheet1.Cells(1, lMaxCols)).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Column + 1
            'clear contents first
            xlSheet1.Range(xlSheet1.Cells(lRow, lCol), xlSheet1.Cells(lMaxRows, lCol)).Value2 = vbNullString

            'keep going until reach last entry or end of spreadsheet
            While Not bDone And lRow < lMaxRows And lBlankRows < 10
                'handle blank cells
                sCellVal = ""
                If xlSheet1.Cells(lRow, 1).value2 Is Nothing Then
                    'count consecutive blanks
                    lBlankRows = lBlankRows + 1
                Else
                    sCellVal = xlSheet1.Cells(lRow, 1).value2
                    lBlankRows = 0 'reset
                End If
                Select Case sCellVal.ToString.Trim.ToLower()
                    'keep track of how far through we are to handle duplicates:
                    'alignment
                    'allow holes
                    'curved
                    'curved (no flip)
                    'feature geom
                    'pref offset
                    'straight
                    'straight (no flip)
                    'Graticule:
                    'offset curved
                    'offset distance
                    'offset horz
                    'regular
                    'repeat
                    'river
                    'spread chars
                    Case "polygons:"
                        bPolygons = True
                    Case "graticule:"
                        If bPoints Then
                            If mxdProps.bPointGACurv Or mxdProps.bPointGACurvNoFlip Or mxdProps.bPointGAStr Or _
                                mxdProps.bPointGAStrNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bLines Then
                            If mxdProps.bLineGACurv Or mxdProps.bLineGACurvNoFlip Or mxdProps.bLineGAStr Or _
                                mxdProps.bLineGAStrNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyGACurv Or mxdProps.bPolyGACurvNoFlip Or mxdProps.bPolyGAStr Or _
                                mxdProps.bPolyGAStrNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                        bGraticule = True
                    Case "lines:"
                        bLines = True
                        bGraticule = False
                    Case "points:"
                        bPoints = True
                        bGraticule = False
                    Case "strategies:"
                        bStrategies = True
                    Case "conflicts:"
                    Case "text symbol:"
                    Case "misc:"
                        'Polys
                    Case "horizontal"
                        If bLines Then
                            If mxdProps.bStreetHorz Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyHorz Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "straight"
                        If bPoints Then
                            If mxdProps.bPointGAStr Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bLines Then
                            If mxdProps.bLineGAStr Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If bGraticule Then
                                If mxdProps.bPolyGAStr Then xlSheet1.Cells(lRow, lCol) = "x"
                            Else
                                If mxdProps.bPolyStr Then xlSheet1.Cells(lRow, lCol) = "x"
                            End If
                        End If
                    Case "curved"
                        If bPoints Then
                            If mxdProps.bPointGACurv Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bLines Then
                            If mxdProps.bLineGACurv Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If bGraticule Then
                                If mxdProps.bPolyGACurv Then xlSheet1.Cells(lRow, lCol) = "x"
                            Else
                                If mxdProps.bPolyCurv Then xlSheet1.Cells(lRow, lCol) = "x"
                            End If
                        End If
                    Case "offset horizontal"
                        If bLines Then
                            If mxdProps.bLineOffHor Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyOffHorz Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "offset curved"
                        If bLines Then
                            If mxdProps.bLineOffCur Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyOffCurv Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "regular"
                        If bLines Then
                            If mxdProps.bLineRegular Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyRegular Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "land parcel"
                        If mxdProps.bPolyParcel Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "river"
                        If bLines Then
                            If mxdProps.bLineRiver Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyRiver Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "boundary"
                        If mxdProps.bPolyBdy Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "single sided"
                        If mxdProps.bPolyBdySingleSided Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "allow holes"
                        If bGraticule Then
                            If mxdProps.bPolyAllowHoles Then xlSheet1.Cells(lRow, lCol) = "x"
                        Else
                            If mxdProps.bPolyBdyAllowHoles Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "on line"
                        If mxdProps.bPolyBdyOnLine Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "try horizontal first"
                        If mxdProps.bPolyTryHorz Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "may place outside"
                        If mxdProps.bPolyMayPlaceOutside Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "offset distance"
                        If bPoints Then
                            If mxdProps.bPointOffDist Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bLines Then
                            If mxdProps.bLineOffDist Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyOffDist Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "preferred offset"
                        If bPoints Then
                            If mxdProps.bPointMaxOffset Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyMaxOffset Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "feature geometry"
                        If bPoints Then
                            If mxdProps.bPointFtrGeom Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bLines Then
                            If mxdProps.bLineFtrGeom Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyFtrGeom Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "straight (no flip)"
                        If bPoints Then
                            If mxdProps.bPointGAStrNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bLines Then
                            If mxdProps.bLineGAStrNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyGAStrNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "curved (no flip)"
                        If bPoints Then
                            If mxdProps.bPointGACurvNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bLines Then
                            If mxdProps.bLineGACurvNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyGACurvNoFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "internal zones"
                        If mxdProps.bPolyIntZones Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "external zones"
                        If mxdProps.bPolyExtZones Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "anchor points"
                        If mxdProps.bPolyAnchor Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "repeat"
                        If bLines Then
                            If mxdProps.bLineRepeat Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolyRepeat Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                    Case "spread chars"
                        If bLines Then
                            If mxdProps.bLineSpread Then xlSheet1.Cells(lRow, lCol) = "x"
                        ElseIf bPolygons Then
                            If mxdProps.bPolySpread Then xlSheet1.Cells(lRow, lCol) = "x"
                        End If
                        'Lines
                    Case "centered horizontal"
                        If mxdProps.bLineCenHor Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "centered straight"
                        If mxdProps.bLineCenStr Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "centered curved"
                        If mxdProps.bLineCenCur Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "centered perpendicular"
                        If mxdProps.bLineCenPer Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "offset straight"
                        If mxdProps.bLineOffStr Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "offset perpendicular"
                        If mxdProps.bLineOffPer Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "street"
                        If mxdProps.bLineStreet Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "reduce leading"
                        If mxdProps.bStreetReduce Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "primary name under"
                        If mxdProps.bStreetPrimary Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "spread words"
                        If mxdProps.bStreetSpread Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "street address"
                        If mxdProps.bLineStreetAdd Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "contour"
                        If mxdProps.bLineContour Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "alignment"
                        If bStrategies Then
                            If mxdProps.bStackC Then AddToCell(xlSheet1.Cells(lRow, lCol), "C", True)
                            If mxdProps.bStackR Then AddToCell(xlSheet1.Cells(lRow, lCol), "R", True)
                            If mxdProps.bStackL Then AddToCell(xlSheet1.Cells(lRow, lCol), "L", True)
                            If mxdProps.bStackLorR Then AddToCell(xlSheet1.Cells(lRow, lCol), "LR", True)
                        ElseIf bLines Then
                            If mxdProps.bContourPage Then AddToCell(xlSheet1.Cells(lRow, lCol), "P")
                            If mxdProps.bContourUphill Then AddToCell(xlSheet1.Cells(lRow, lCol), "U")
                        End If
                    Case "ladders"
                        If mxdProps.bContourLadder Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "secondary offset"
                        If mxdProps.bLineSecOff Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "constraint"
                        If mxdProps.bConstrainAbove Then AddToCell(xlSheet1.Cells(lRow, lCol), "A", True)
                        If mxdProps.bConstrainBelow Then AddToCell(xlSheet1.Cells(lRow, lCol), "B", True)
                        If mxdProps.bConstrainLeft Then AddToCell(xlSheet1.Cells(lRow, lCol), "L", True)
                        If mxdProps.bConstrainRight Then AddToCell(xlSheet1.Cells(lRow, lCol), "R", True)
                        If mxdProps.bNoConstraint Then AddToCell(xlSheet1.Cells(lRow, lCol), "N", True)
                    Case "position"
                        If mxdProps.bLineBestAlong Then AddToCell(xlSheet1.Cells(lRow, lCol), "B", True)
                        If mxdProps.bLineBeforeStart Then AddToCell(xlSheet1.Cells(lRow, lCol), "BS", True)
                        If mxdProps.bLineAfterEnd Then AddToCell(xlSheet1.Cells(lRow, lCol), "AE", True)
                        If mxdProps.bLineFromStart Then AddToCell(xlSheet1.Cells(lRow, lCol), "AFS", True)
                        If mxdProps.bLineFromEnd Then AddToCell(xlSheet1.Cells(lRow, lCol), "AFE", True)
                    Case "straddlacking"
                        If mxdProps.bStraddlacking Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "align to direction"
                        If mxdProps.bLineDirection Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "label near border"
                        If mxdProps.bLabelNearBorder Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "label near junction"
                        If mxdProps.bLabelNearJunction Then xlSheet1.Cells(lRow, lCol) = "x"
                        'Points
                    Case "fixed"
                        If mxdProps.bPointFixed Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "may shift"
                        If mxdProps.bMayShift Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "best"
                        If mxdProps.bPointBest Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "zones"
                        If mxdProps.bPointZones Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "altered zones"
                        If mxdProps.bAlteredZones Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "symbol outline"
                        If mxdProps.bSymbolOutline Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "rotation:"
                        If mxdProps.bPointRotation Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "additional angle"
                        If mxdProps.bPointRotAngle Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "may flip"
                        If mxdProps.bPointRotFlip Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "multipoint"
                        If mxdProps.bMultipoint Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "whitespace"
                        If mxdProps.bWhitespace Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "line breaks"
                        If mxdProps.bLinebreaks Then xlSheet1.Cells(lRow, lCol) = "x"
                        'Strategies
                    Case "stack:"
                        If mxdProps.bStack Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "stacking chars"
                        For i = 0 To mxdProps.lSeparators - 1
                            If InStr(1, mxdProps.sSeparators(i), " ") = 0 And InStr(1, mxdProps.sSeparators(i), ",") = 0 _
                            And InStr(1, mxdProps.sSeparators(i), "-") = 0 Then _
                            AddToCell(xlSheet1.Cells(lRow, lCol), mxdProps.sSeparators(i))
                        Next
                    Case "limits"
                        If mxdProps.bMaxLines Then AddToCell(xlSheet1.Cells(lRow, lCol), "1")
                        If mxdProps.bMinChars Then AddToCell(xlSheet1.Cells(lRow, lCol), "2")
                        If mxdProps.bMaxChars Then AddToCell(xlSheet1.Cells(lRow, lCol), "3")
                    Case "overrun (value)"
                        If mxdProps.bOverrun Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "overrun (asymmetric)"
                        If mxdProps.bAsymmetric Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "font reduction"
                        If mxdProps.bFontReduction Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "font compression"
                        If mxdProps.bCompression Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "abbreviation"
                        If mxdProps.bAbbreviation Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "translation"
                        If mxdProps.bAbbreviation And mxdProps.bDictionaryTranslation Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "keyword"
                        If mxdProps.bAbbreviation And mxdProps.bDictionaryKeyword Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "ending"
                        If mxdProps.bAbbreviation And mxdProps.bDictionaryEnding Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "truncation"
                        If mxdProps.bTruncation Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "length"
                        If mxdProps.bTruncationLength Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "marker"
                        If mxdProps.bTruncationMarker Then
                            For i = 0 To mxdProps.lTruncMarker - 1
                                If InStr(1, mxdProps.sTruncMarker(i), ".") = 0 Then _
                                AddToCell(xlSheet1.Cells(lRow, lCol), mxdProps.sTruncMarker(i))
                            Next
                        End If
                    Case "preferred chars"
                        If mxdProps.bTruncationChars Then
                            For i = 0 To mxdProps.lTruncChars - 1
                                If InStr(1, mxdProps.sTruncChars(i), "aeiou") = 0 Then _
                                AddToCell(xlSheet1.Cells(lRow, lCol), mxdProps.sTruncChars(i))
                            Next
                        End If
                    Case "key numbering"
                        If mxdProps.bKeyNumbering Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "strategy order"
                        If mxdProps.bStrategyPriority Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "minimum size"
                        If mxdProps.bMinSize Then xlSheet1.Cells(lRow, lCol) = "x"
                        'Conflicts
                    Case "weights"
                        If mxdProps.bWeights Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "background labels"
                        If mxdProps.bBackground Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "remove duplicates"
                        If mxdProps.bRemoveDup Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "label buffer"
                        If mxdProps.bLabelBuffer Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "hard constraint"
                        If mxdProps.bHardConstraint Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "never remove"
                        If mxdProps.bNeverRemove Then xlSheet1.Cells(lRow, lCol) = "x"
                        'Text Symbol
                    Case "x/y offset"
                        If mxdProps.bXYOffset Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "right to left"
                        If mxdProps.bRighttoLeft Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "text position"
                        If mxdProps.bTextPosition Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "text case"
                        If mxdProps.bTextCase Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "char spacing"
                        If mxdProps.bCharSpacing Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "leading"
                        If mxdProps.bLeading Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "char width"
                        If mxdProps.bCharWidth Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "word spacing"
                        If mxdProps.bWordSpacing Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "kerning off"
                        If mxdProps.bKerningOff Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "fill symbol"
                        If mxdProps.bFillSymbol Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "text background:"
                        If mxdProps.bTextBackground Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "balloon callout"
                        If mxdProps.bBalloonCallout Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "line callout"
                        If mxdProps.bLineCallout Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "marker text background"
                        If mxdProps.bMarkerTextBkg Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "scale to fit text"
                        If mxdProps.bScaletoFit Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "simple line callout"
                        If mxdProps.bSimpleLineCallout Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "shadow"
                        If mxdProps.bShadow Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "halo"
                        If mxdProps.bHalo Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "cjk"
                        If mxdProps.bCJK Then xlSheet1.Cells(lRow, lCol) = "x"
                        'Misc
                    Case "layer definition query"
                        If mxdProps.bLayerDefQuery Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "sql query"
                        If mxdProps.bSQL Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "label expression"
                        If mxdProps.bLabelExpression Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "tags bse"
                        If mxdProps.bBaseTag Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "tags other"
                        If mxdProps.bTags Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "scale ranges"
                        If mxdProps.bScaleRanges Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "rotate labels with dataframe"
                        If mxdProps.bRotateWithDataFrame Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "inverted label tolerance"
                        xlSheet1.Cells(lRow, lCol) = mxdProps.sInvertedLabTol
                    Case "line connection"
                        If mxdProps.bUnambiguous Then AddToCell(xlSheet1.Cells(lRow, lCol), "U")
                        If mxdProps.bMinimize Then AddToCell(xlSheet1.Cells(lRow, lCol), "M")
                    Case "allow border overlap"
                        If mxdProps.bAllowOverlap Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "label largest only"
                        If mxdProps.bLargestOnly Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "data frame rotation"
                        If mxdProps.bFrameRotation Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "label priority ranking"
                        If mxdProps.lLabelClassCount > 1 Then If mxdProps.bLabelPriority Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "graphic barriers/anno"
                        If mxdProps.lAnnoLayers > 0 Or mxdProps.lBarriers > 0 Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "fast placement"
                        If mxdProps.bFast Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "map units"
                        For i = 0 To mxdProps.lMapCount - 1
                            If StrComp(mxdProps.sMapUnits(i), "meters", CompareMethod.Text) = 0 Then
                                AddToCell(xlSheet1.Cells(lRow, lCol), "M", True)
                            ElseIf StrComp(mxdProps.sMapUnits(i), "decimal degrees", CompareMethod.Text) = 0 Then
                                AddToCell(xlSheet1.Cells(lRow, lCol), "DD", True)
                            ElseIf StrComp(mxdProps.sMapUnits(i), "feet", CompareMethod.Text) = 0 Then
                                AddToCell(xlSheet1.Cells(lRow, lCol), "F", True)
                            Else
                                AddToCell(xlSheet1.Cells(lRow, lCol), "Other", True)
                            End If
                        Next
                    Case "spatial reference"
                        If mxdProps.bGeographic Then AddToCell(xlSheet1.Cells(lRow, lCol), "G", True)
                        If mxdProps.bProjected Then AddToCell(xlSheet1.Cells(lRow, lCol), "P", True)
                    Case "reference scale"
                        If mxdProps.bRefScale Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "automatic extent"
                        If mxdProps.bAutoExtent Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "fixed extent"
                        If mxdProps.bFixedExtent Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "fixed scale"
                        If mxdProps.bFixedScale Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "clip extent"
                        If mxdProps.bClipExtent Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "clip to shape"
                        If mxdProps.bClipToShape Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "exclude layers"
                        If mxdProps.bExcludeLayers Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "multiple dataframes"
                        If mxdProps.lMapCount > 1 Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "data view"
                        If Not mxdProps.bLayoutView Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "layout view"
                        If mxdProps.bLayoutView Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "shapefiles"
                        If mxdProps.bSHP Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "personal gdb"
                        If mxdProps.bPGDB Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "file gdb"
                        If mxdProps.bFGDB Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "sde"
                        If mxdProps.bSDE Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "qualified names"
                        If mxdProps.bQualifiedNames Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "html entities"
                        If mxdProps.bHTMLEnt Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "coded value domain"
                        If mxdProps.bCodedValueDomain Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "unplaced labels"
                        If mxdProps.bDrawUnplaced Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "absolute paths"
                        If mxdProps.bAbsPaths Then xlSheet1.Cells(lRow, lCol) = "x"
                        'symbols
                    Case "bar chart"
                        If mxdProps.bBarChart Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "pie chart"
                        If mxdProps.bPieChart Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "stacked chart"
                        If mxdProps.bStackedChart Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "simple fill"
                        If mxdProps.bSimpleFill Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "gradient fill"
                        If mxdProps.bGradientFill Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "marker fill"
                        If mxdProps.bMarkerFill Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "picture fill"
                        If mxdProps.bPictureFill Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "line fill"
                        If mxdProps.bLineFill Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "fixed size"
                        If mxdProps.bFixedSize Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "leaders"
                        If mxdProps.bChartLeaders Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "3d chart"
                        If mxdProps.b3DChart Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "chart overlap"
                        If mxdProps.bChartOverlap Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "bar orientation"
                        If mxdProps.bBarOrient Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "column orientation"
                        If mxdProps.bColumnOrient Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "arithmetic orientation"
                        If mxdProps.bArithOrient Then xlSheet1.Cells(lRow, lCol) = "x"
                    Case "geographic orientation"
                        If mxdProps.bGeogOrient Then xlSheet1.Cells(lRow, lCol) = "x"

                        'end marker
                        'Case "volume"
                        '  bDone = True
                End Select
                lRow = lRow + 1
            End While

            xlSheet1.Cells(1, lCol) = Replace(Right(sMxdName, Len(sMxdName) - InStrRev(sMxdName, "\")), ".mxd", "")

        Catch e As Exception
            sw.WriteLine(e.ToString)
        Finally
            'save and close
            xlApp.ActiveWorkbook.Close(True)
            xlApp.Quit()
            releaseObject(xlSheet1)
            releaseObject(xlWbook)
            releaseObject(xlApp)
        End Try
    End Sub

    Private Sub releaseObject(ByRef obj As Object)
        If Not obj Is Nothing Then
            Try
                Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
                MessageBox.Show("Unable to release the Object " + ex.ToString())
            Finally
                GC.Collect()
            End Try
        End If
    End Sub

    Private Sub AddToCell(ByRef xlCell As Microsoft.Office.Interop.Excel.Range, ByVal str As String, _
                          Optional ByVal bSpace As Boolean = False)

        Dim sTmp As String = vbNullString
        If Not xlCell.Value2 = Nothing Then
            sTmp = xlCell.Value2.ToString
            If bSpace Then sTmp = sTmp & " "
        End If
        sTmp = sTmp & str
        xlCell.Value2 = sTmp

    End Sub

    Public Sub DeleteDirectory(ByVal sDirectory As String)

        If Directory.Exists(sDirectory) Then
            Try
                MakeDirectoryWriteable(sDirectory)
                Directory.Delete(sDirectory, True)
            Catch ex As Exception
                Console.WriteLine("DeleteDirectory error: Could not delete folder [" & sDirectory & "]. " & ex.Message)
            End Try
        End If

    End Sub

    Public Sub MakeDirectoryWriteable(ByVal sDirectoryName As String)
        'Make folder and contents of folder writeable if it exists
        If Directory.Exists(sDirectoryName) Then
            Dim pDirInfo As DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(sDirectoryName)
            pDirInfo.Attributes = FileAttributes.Normal
        End If
    End Sub

End Module