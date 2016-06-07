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

Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.GISClient
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports System.Text.RegularExpressions


Module ModFunctions

    Public Sub GetLayerProps(ByRef pLayer As ILayer, ByRef lTabLevel As Integer, Optional ByRef bSymbolLevels As Boolean = False)
        sw.Flush()
        Dim pSubLayer As ILayer
        Dim pCompLayer As ICompositeLayer
        Dim pGraphicsLayer As IGraphicsLayer
        Dim pFL As IFeatureLayer2
        Dim pGeoFL As IGeoFeatureLayer
        Dim pRasterLayer As IRasterLayer
        Dim pDataLayer As IDataLayer2
        Dim pDatasetName As IDatasetName
        Dim pWSName As IWorkspaceName
        Dim pMapServerLayer As IMapServerLayer
        Dim pIAGSName As IAGSServerObjectName3 = Nothing
        Dim pBasemapLayer As IBasemapLayer
        Dim pFLayerDef As IFeatureLayerDefinition
        Dim pLayerFX As ILayerEffects
        Dim pAnnotateLPColl As IAnnotateLayerPropertiesCollection
        Dim pAnnotateLP As IAnnotateLayerProperties = Nothing
        Dim pLabelEngineLayerProperties As ILabelEngineLayerProperties2
        Dim pMpxOpLProps As IMaplexOverposterLayerProperties
        Dim pBasicOpLProps As IBasicOverposterLayerProperties4
        Dim pAnnoClassExt As IAnnotationClassExtension
        Dim pAnnoSubLayer As IAnnotationSublayer
        Dim pSymSubst As ISymbolSubstitution
        Dim pSymColl As ISymbolCollection2
        Dim pTextSym As ITextSymbol
        Dim pSymbolLevels As ISymbolLevels
        Dim sTmp As String = vbNullString
        Dim sDoc As String = vbNullString
        Dim sMap As String = vbNullString
        Dim lLabelClass As Integer, iSelFtrs As Integer, lGLayer As Integer
        Dim i As Integer, j As Integer

        'layer name, scale range and visibility
        sw.WriteLine(InsertTabs(lTabLevel) & "Name: " & pLayer.Name & ", Scale range: " & pLayer.MinimumScale & " - " & pLayer.MaximumScale)
        sw.WriteLine(InsertTabs(lTabLevel) & "Visible: " & pLayer.Visible)

        'get data source
        If TypeOf pLayer Is IDataLayer Then
            Try
                pDataLayer = pLayer
                If Not (TypeOf pLayer Is IBasemapLayer Or _
                        TypeOf pLayer Is IMapServerLayer Or _
                        TypeOf pLayer Is IMapServerRESTLayer Or _
                        TypeOf pLayer Is IImageServerLayer) Then
                    pDatasetName = pDataLayer.DataSourceName
                    pWSName = pDatasetName.WorkspaceName
                    mxdProps.sDataSources(mxdProps.lDataSources) = pWSName.PathName
                    sw.WriteLine(InsertTabs(lTabLevel) & "Data source: " & mxdProps.sDataSources(mxdProps.lDataSources))
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Dataset name: " & pDatasetName.Name)
                    AddIfUnique(mxdProps.lDataSources, mxdProps.sDataSources, ARRAY_SIZE)

                    'find datasource type
                    Dim pWksF As WorkspaceFactory
                    Dim bDone As Boolean = False
                    'personal gdb
                    If Not bDone And Not mxdProps.bPGDB Then
                        pWksF = New AccessWorkspaceFactory
                        If pWksF.IsWorkspace(pWSName.PathName) Then
                            mxdProps.bPGDB = True
                            bDone = True
                        End If
                    End If
                    'file gdb
                    If Not bDone And Not mxdProps.bFGDB Then
                        pWksF = New FileGDBWorkspaceFactory
                        If pWksF.IsWorkspace(pWSName.PathName) Then
                            mxdProps.bFGDB = True
                            bDone = True
                        End If
                    End If
                    'shapefile
                    If Not bDone And Not mxdProps.bSHP Then
                        pWksF = New ShapefileWorkspaceFactory
                        If pWksF.IsWorkspace(pWSName.PathName) Then
                            mxdProps.bSHP = True
                            bDone = True
                        End If
                    End If
                    'arcinfo coverage
                    If Not bDone And Not mxdProps.bCoverage Then
                        pWksF = New ArcInfoWorkspaceFactory
                        If pWksF.IsWorkspace(pWSName.PathName) Then
                            mxdProps.bCoverage = True
                            bDone = True
                        End If
                    End If
                    'pc coverage
                    If Not bDone And Not mxdProps.bCoverage Then
                        pWksF = New PCCoverageWorkspaceFactory
                        If pWksF.IsWorkspace(pWSName.PathName) Then
                            mxdProps.bCoverage = True
                            bDone = True
                        End If
                    End If
                    'sde
                    If Not bDone And Not mxdProps.bSDE Then
                        pWksF = New SdeWorkspaceFactory
                        If pWksF.IsWorkspace(pWSName.PathName) Then
                            mxdProps.bSDE = True
                            bDone = True
                        End If
                    End If
                End If 'not server layer
            Catch ex As Exception
                sw.WriteLine(InsertTabs(lTabLevel) & "Error: " & ex.ToString)
            End Try
        End If 'is DataLayer

        'get layer type
        If TypeOf pLayer Is IGeoFeatureLayer Then
            pFL = pLayer
            sw.WriteLine(InsertTabs(lTabLevel) & "Feature layer")
            pLayerFX = pLayer
            If pLayerFX.Transparency > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Transparency: " & pLayerFX.Transparency & "%")
            If pLayer.Visible Or bAllLayers Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Geometry: " & GetGeomType((pFL.ShapeType)))
                sw.WriteLine(InsertTabs(lTabLevel) & "Scale symbols: " & pFL.ScaleSymbols)
            End If
            If pFL.ShapeType = esriGeometryType.esriGeometryMultiPatch Then mxdProps.bMultipatch = True
            If pFL.ShapeType = esriGeometryType.esriGeometryMultipoint Then mxdProps.bMultipoint = True
            iSelFtrs = GetSelectedFeatures(pLayer, lTabLevel)
            If iSelFtrs > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Selected features: " & iSelFtrs)

            'if not visible, do nothing else
            If pLayer.Visible Or bAllLayers Then
                'Layer def query
                pFLayerDef = pLayer
                If Len(pFLayerDef.DefinitionExpression) > 0 Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Definition Query: " & pFLayerDef.DefinitionExpression)
                    If IsQualifiedName(pFLayerDef.DefinitionExpression) Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Qualified name!")
                    mxdProps.bLayerDefQuery = True
                End If

                pGeoFL = pLayer
                If pGeoFL.ShowTips Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Show map tips, display field: " & pGeoFL.DisplayField)
                    If IsQualifiedName(pGeoFL.DisplayField) Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Qualified name!")
                End If
                If bReadSymbols Then
                    If TypeOf pGeoFL.Renderer Is ILevelRenderer Then
                        pSymbolLevels = pGeoFL
                        If pSymbolLevels.UseSymbolLevels Then
                            bSymbolLevels = True
                            sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels")
                        Else
                            If bSymbolLevels Then sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels (inherited)")
                        End If
                    End If
                    If Not pGeoFL.Renderer Is Nothing Then GetRendererProps(pGeoFL.Renderer, lTabLevel + 1, bSymbolLevels)
                End If
                If bReadLabels Then
                    'label class visibility
                    If pGeoFL.DisplayAnnotation Or bAllLayers Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Display label classes: " & pGeoFL.DisplayAnnotation)
                        GetAnnoProps(pLayer, pGeoFL.AnnotationProperties, pGeoFL.AnnotationPropertiesID, bLyrFile, lTabLevel)

                    Else 'display or show all: feature class
                        'include weights even if not visible
                        sw.WriteLine(InsertTabs(lTabLevel) & "Display label classes: " & pGeoFL.DisplayAnnotation & " - showing weighted label classes only")
                        pAnnotateLPColl = pGeoFL.AnnotationProperties
                        If pAnnotateLPColl Is Nothing Then Return

                        For lLabelClass = 0 To pAnnotateLPColl.Count - 1
                            pAnnotateLPColl.QueryItem(lLabelClass, pAnnotateLP, Nothing, Nothing)
                            pLabelEngineLayerProperties = pAnnotateLP
                            If mxdProps.bMapIsMLE Then
                                pMpxOpLProps = pLabelEngineLayerProperties.OverposterLayerProperties
                                If pMpxOpLProps.FeatureWeight Or _
                                    (pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon And _
                                    pMpxOpLProps.PolygonBoundaryWeight) Then
                                    sw.WriteLine(InsertTabs(lTabLevel) & "Label class " & pAnnotateLP.Class & ":")
                                    If pMpxOpLProps.FeatureWeight Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Weight: " & pMpxOpLProps.FeatureWeight)
                                    If (pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon And _
                                        pMpxOpLProps.PolygonBoundaryWeight) Then sw.WriteLine(InsertTabs(lTabLevel + 1) & _
                                        "Boundary weight: " & pMpxOpLProps.PolygonBoundaryWeight)
                                    mxdProps.bWeights = True
                                End If 'weight
                            End If
                            If mxdProps.bMapIsSLE Then
                                pBasicOpLProps = pLabelEngineLayerProperties.BasicOverposterLayerProperties
                                If pBasicOpLProps.FeatureWeight Or pBasicOpLProps.LabelWeight Then
                                    sw.WriteLine(InsertTabs(lTabLevel) & "Label class " & pAnnotateLP.Class & ":")
                                    If pBasicOpLProps.FeatureWeight Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Feature weight: " & pBasicOpLProps.FeatureWeight)
                                    If pBasicOpLProps.LabelWeight Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Label weight: " & pBasicOpLProps.LabelWeight)
                                    If pBasicOpLProps.LabelWeight <> esriBasicOverposterWeight.esriHighWeight Or pBasicOpLProps.FeatureWeight Then mxdProps.bWeights = True
                                End If 'weight
                            End If 'mle
                        Next  'label class
                    End If 'display or show all: label class
                End If 'ReadSymbols
            End If 'display or show all: feature class
            '
            '********************
        ElseIf TypeOf pLayer Is IDimensionLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Dimension layer")

            '********************
        ElseIf TypeOf pLayer Is IFDOGraphicsLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Anno layer")
            If bReadLabels Then
                pLayerFX = pLayer
                If pLayerFX.Transparency > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Transparency: " & pLayerFX.Transparency & "%")
                iSelFtrs = GetSelectedFeatures(pLayer, lTabLevel)
                If iSelFtrs > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Selected features: " & iSelFtrs)
                mxdProps.lAnnoLayers = mxdProps.lAnnoLayers + 1
                'Symbol substitution
                pSymSubst = pLayer
                Select Case pSymSubst.SubstituteType
                    Case esriSymbolSubstituteType.esriSymbolSubstituteColor
                        sw.WriteLine(InsertTabs(lTabLevel + 1) & "Symbol substitution color (RGB): " & GetRGB(pSymSubst.MassColor))
                        sw.WriteLine(InsertTabs(lTabLevel + 1) & "Symbol substitution color (CMYK): " & GetCMYK(pSymSubst.MassColor))
                    Case esriSymbolSubstituteType.esriSymbolSubstituteNone
                        'do nothing
                    Case Else
                        If pSymSubst.SubstituteType = esriSymbolSubstituteType.esriSymbolSubstituteIndividualDominant Then
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Symbol substitution (symbol precedence)")
                        ElseIf pSymSubst.SubstituteType = esriSymbolSubstituteType.esriSymbolSubstituteIndividualSubordinate Then
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Symbol substitution (override precedence)")
                        End If
                        If Not pSymSubst.InlineColor Is Nothing Then
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Inline color (RGB): " & GetRGB(pSymSubst.InlineColor))
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Inline color (CMYK): " & GetCMYK(pSymSubst.InlineColor))
                        End If
                        Try
                            pSymColl = pSymSubst.SubstituteSymbolCollection
                            For j = 0 To pSymColl.Count - 1
                                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Symbol " & j & ": ")
                                pTextSym = pSymColl.Symbol(j)
                                GetTextSymbolProps(pTextSym, lTabLevel + 2)
                            Next
                        Catch ex As Exception
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Error: " & ex.ToString)
                        End Try
                End Select
                pCompLayer = pLayer
                For i = 0 To pCompLayer.Count - 1
                    pSubLayer = pCompLayer.Layer(i)
                    'sw.WriteLine(InsertTabs(lTabLevel + 1) & "Anno class: " & pSubLayer.Name & ", Scale range: " & pSubLayer.MinimumScale & " - " & pSubLayer.MaximumScale)
                    'sw.WriteLine(InsertTabs(lTabLevel + 2) & "Visible: " & pSubLayer.Visible)

                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Anno layer " & i + 1 & "\" & pCompLayer.Count)
                    GetLayerProps(pSubLayer, lTabLevel + 2)
                Next
            End If 'read labels

            '********************
        ElseIf TypeOf pLayer Is IAnnotationSublayer Then
            pAnnoSubLayer = pLayer
            sw.WriteLine(InsertTabs(lTabLevel) & "Annotation Class ID: " & pAnnoSubLayer.AnnotationClassID)
            If bReadLabels Then
                ' Get the feature class from the feature layer.
                Dim featureLayer As IFeatureLayer = CType(pAnnoSubLayer.Parent, IFeatureLayer)
                Dim featureClass As IFeatureClass = featureLayer.FeatureClass
                pAnnoClassExt = featureClass.Extension
                GetAnnoProps(pLayer, pAnnoClassExt.AnnoProperties, Nothing, False, lTabLevel)
            End If 'read labels

            '********************
        ElseIf TypeOf pLayer Is ICompositeGraphicsLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Composite Graphics layer")
            If bReadLabels Then
                pCompLayer = pLayer
                For i = 0 To pCompLayer.Count - 1
                    pSubLayer = pCompLayer.Layer(i)
                    pGraphicsLayer = pSubLayer
                    sTmp = "Anno class: " & pSubLayer.Name
                    If Not pGraphicsLayer.AssociatedLayer Is Nothing Then _
                        sTmp = sTmp & ", Associated layer: " & pGraphicsLayer.AssociatedLayer.Name
                    If pSubLayer.MaximumScale Or pSubLayer.MinimumScale Then _
                      sTmp = sTmp & ", Scale range: " & pSubLayer.MinimumScale & " - " & pSubLayer.MaximumScale
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & sTmp)
                    sw.WriteLine(InsertTabs(lTabLevel + 2) & "Visible: " & pSubLayer.Visible)
                Next
            End If 'read labels

            '********************
        ElseIf TypeOf pLayer Is IGraphicsLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Graphics layer")

            '********************
        ElseIf TypeOf pLayer Is ICoverageAnnotationLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Coverage anno layer")
            If bReadLabels Then
                pLayerFX = pLayer
                If pLayerFX.Transparency > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Transparency: " & pLayerFX.Transparency & "%")
            End If 'read labels

            '********************
        ElseIf TypeOf pLayer Is IMapServerRESTLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Map server REST layer")
            'TODO

        ElseIf TypeOf pLayer Is IMapServerLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Map server layer")
            pLayerFX = pLayer
            If pLayerFX.Transparency > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Transparency: " & pLayerFX.Transparency & "%")
            If bReadLabels Then
                Try
                    pMapServerLayer = pLayer
                    pMapServerLayer.GetConnectionInfo(pIAGSName, sDoc, sMap)
                    If Not pIAGSName Is Nothing Then
                        If pIAGSName.Name() <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Server: " & pIAGSName.Name())
                        If pIAGSName.Description() <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Description: " & pIAGSName.Description())
                        If pIAGSName.URL() <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "URL: " & pIAGSName.URL())
                    End If
                    If sDoc <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Document location: " & sDoc)
                    If sMap <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Map name: " & sMap)
                Catch ex As Exception
                    sw.WriteLine(InsertTabs(lTabLevel) & "Error: could not get connection info. " & ex.Message)
                End Try
                'Exit Sub
            End If 'read labels

            '********************
        ElseIf TypeOf pLayer Is IBasemapLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Basemap layer")
            If bReadLabels Then
                pBasemapLayer = pLayer
                If pBasemapLayer.CanDraw Then sw.WriteLine(InsertTabs(lTabLevel) & "Can draw")
            End If 'read labels

            '********************
        ElseIf TypeOf pLayer Is IImageServerLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Image server layer")
            pLayerFX = pLayer
            If pLayerFX.Transparency > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Transparency: " & pLayerFX.Transparency & "%")
            If bReadSymbols Then
                Dim pImageServerLayer As IImageServerLayer = pLayer
                If TypeOf pImageServerLayer.Renderer Is ILevelRenderer Then
                    pSymbolLevels = pImageServerLayer
                    If pSymbolLevels.UseSymbolLevels Then
                        bSymbolLevels = True
                        sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels")
                    Else
                        If bSymbolLevels Then sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels (inherited)")
                    End If

                    If Not pImageServerLayer.Renderer Is Nothing Then GetRendererProps(pImageServerLayer.Renderer, lTabLevel + 1, bSymbolLevels)
                End If
            End If 'read symbols
            '********************
        ElseIf TypeOf pLayer Is IRasterLayer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Raster layer")
            pRasterLayer = pLayer

            sw.WriteLine(InsertTabs(lTabLevel) & "Columns and Rows: " & pRasterLayer.ColumnCount & ", " & pRasterLayer.RowCount)
            sw.WriteLine(InsertTabs(lTabLevel) & "Band count: " & pRasterLayer.BandCount)
            ' can't get pyramid info if data link is broken
            Try
                sw.WriteLine(InsertTabs(lTabLevel) & "Pyramids: " & pRasterLayer.PyramidPresent)
            Catch ex As Exception
                sw.WriteLine(InsertTabs(lTabLevel) & "Error: could not get pyramid info. " & ex.Message)
            End Try
            If Not pRasterLayer.Raster Is Nothing Then
                Dim pRas As IRaster
                pRas = pRasterLayer.Raster
                Select Case pRas.ResampleMethod
                    Case rstResamplingTypes.RSP_Average
                        sw.WriteLine(InsertTabs(lTabLevel) & "Resampling type: Average")
                    Case rstResamplingTypes.RSP_CubicConvolution
                        sw.WriteLine(InsertTabs(lTabLevel) & "Resampling type: Cubic convolution")
                    Case rstResamplingTypes.RSP_BilinearInterpolation
                        sw.WriteLine(InsertTabs(lTabLevel) & "Resampling type: Bilinear interpolation")
                    Case rstResamplingTypes.RSP_BilinearInterpolationPlus
                        sw.WriteLine(InsertTabs(lTabLevel) & "Resampling type: Bilinear interpolation plus")
                    Case rstResamplingTypes.RSP_Majority
                        sw.WriteLine(InsertTabs(lTabLevel) & "Resampling type: Majority")
                    Case rstResamplingTypes.RSP_NearestNeighbor
                        sw.WriteLine(InsertTabs(lTabLevel) & "Resampling type: Nearest neighbour")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Resampling type: ************ TODO ************")
                End Select
            End If
            If bReadSymbols Then
                If TypeOf pRasterLayer.Renderer Is ILevelRenderer Then
                    pSymbolLevels = pRasterLayer
                    If pSymbolLevels.UseSymbolLevels Then
                        bSymbolLevels = True
                        sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels")
                    Else
                        If bSymbolLevels Then sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels (inherited)")
                    End If
                End If
                If Not pRasterLayer.Renderer Is Nothing Then GetRasterRendererProps(pRasterLayer.Renderer, lTabLevel + 1, bSymbolLevels)
            End If

            '********************
        ElseIf TypeOf pLayer Is IGroupLayer Then
            'get sublayers
            pCompLayer = pLayer
            sw.WriteLine(InsertTabs(lTabLevel) & "Group layer, number of sublayers: " & pCompLayer.Count)
            Try
                pLayerFX = pLayer
                If pLayerFX.Transparency > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Transparency: " & pLayerFX.Transparency & "%")
            Catch ex As Exception
                'group layer does not support transparency
            End Try
            If bReadSymbols Then
                pSymbolLevels = pCompLayer
                If pSymbolLevels.UseSymbolLevels Then
                    bSymbolLevels = True
                    sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels")
                Else
                    If bSymbolLevels Then sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol levels (inherited)")
                End If
            End If
            If pLayer.Visible Or bAllLayers Then
                For lGLayer = 0 To pCompLayer.Count - 1
                    sw.WriteLine(vbCrLf & InsertTabs(lTabLevel) & pLayer.Name & " layer " & lGLayer + 1 & "/" & pCompLayer.Count)
                    GetLayerProps(pCompLayer.Layer(lGLayer), lTabLevel + 1)
                Next lGLayer
            End If

        Else
            sw.WriteLine(InsertTabs(lTabLevel) & "*** Unknown layer type ***")
        End If 'layer type
    End Sub

    Sub GetRasterRendererProps(ByRef pRR As IRasterRenderer, ByRef lTabLevel As Integer, ByRef bSymbolLevels As Boolean)
        sw.Flush()
        'Dim sTmp As String
        Dim i As Integer
        Dim pClassCRRend As IRasterClassifyColorRampRenderer
        '? Dim pColormapRend As IRasterColormapRenderer
        Dim pDiscreteColRend As IRasterDiscreteColorRenderer
        Dim pRGBRend As IRasterRGBRenderer2
        Dim pStretchCRRend As IRasterStretchColorRampRenderer
        Dim pRasterStretch2 As IRasterStretch2
        Dim pRasterStretch3 As IRasterStretch3
        Dim pUVR As IRasterUniqueValueRenderer
        If TypeOf pRR Is IRasterClassifyColorRampRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Raster Classify Color Ramp Renderer")
            mxdProps.bRasterClassify = True
            pClassCRRend = pRR
            sw.WriteLine(InsertTabs(lTabLevel) & "Class count: " & pClassCRRend.ClassCount)
            sw.WriteLine(InsertTabs(lTabLevel) & "Class field: " & pClassCRRend.ClassField)
            For i = 0 To pClassCRRend.ClassCount - 1
                If pClassCRRend.Description(i) <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Description: " & pClassCRRend.Description(i))
                If pClassCRRend.Label(i) <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Label: " & pClassCRRend.Label(i))
                GetSymbolProps(pClassCRRend.Symbol(i), lTabLevel + 1, bSymbolLevels)
            Next
        ElseIf TypeOf pRR Is IRasterDiscreteColorRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Raster Discrete Color Renderer")
            mxdProps.bRasterDiscrete = True
            pDiscreteColRend = pRR
            sw.WriteLine(InsertTabs(lTabLevel) & "Number of colors: " & pDiscreteColRend.NumColors)
        ElseIf TypeOf pRR Is IRasterRGBRenderer Or TypeOf pRR Is IRasterRGBRenderer2 Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Raster RGB Renderer")
            mxdProps.bRasterRGB = True
            pRGBRend = pRR
            sw.WriteLine(InsertTabs(lTabLevel) & "Red band index: " & pRGBRend.RedBandIndex & ", use: " & pRGBRend.UseRedBand)
            sw.WriteLine(InsertTabs(lTabLevel) & "Green band index: " & pRGBRend.GreenBandIndex & ", use: " & pRGBRend.UseGreenBand)
            sw.WriteLine(InsertTabs(lTabLevel) & "Blue band index: " & pRGBRend.BlueBandIndex & ", use: " & pRGBRend.UseBlueBand)
            sw.WriteLine(InsertTabs(lTabLevel) & "Alpha band index: " & pRGBRend.AlphaBandIndex & ", use: " & pRGBRend.UseAlphaBand)
            ' TODO raster stretch is a coclass so make a function to reuse stretch stuff below
        ElseIf TypeOf pRR Is IRasterStretchColorRampRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Raster Stretch Color Ramp Renderer")
            mxdProps.bRasterStretch = True
            pStretchCRRend = pRR
            sw.WriteLine(InsertTabs(lTabLevel) & "Band index: " & pStretchCRRend.BandIndex)
            If pStretchCRRend.ColorScheme <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Color scheme: " & pStretchCRRend.ColorScheme)
            If pStretchCRRend.LabelHigh <> "" Or pStretchCRRend.LabelMedium <> "" Or pStretchCRRend.LabelLow <> "" Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Labels:")
            End If
            If pStretchCRRend.LabelHigh <> "" Then sw.WriteLine(InsertTabs(lTabLevel + 1) & pStretchCRRend.LabelHigh)
            If pStretchCRRend.LabelMedium <> "" Then sw.WriteLine(InsertTabs(lTabLevel + 1) & pStretchCRRend.LabelMedium)
            If pStretchCRRend.LabelLow <> "" Then sw.WriteLine(InsertTabs(lTabLevel + 1) & pStretchCRRend.LabelLow)
            GetColorRampProps(pStretchCRRend.ColorRamp, lTabLevel, bSymbolLevels)
            pRasterStretch2 = pStretchCRRend
            pRasterStretch3 = pStretchCRRend
            If pRasterStretch2.Background Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Background color (RGB): " & GetRGB(pRasterStretch2.BackgroundColor))
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Background color (CMYK): " & GetCMYK(pRasterStretch2.BackgroundColor))
            End If
            Select Case pRasterStretch2.StretchType
                Case esriRasterStretchTypesEnum.esriRasterStretch_StandardDeviations
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Stretch type: Standard deviations")
                Case esriRasterStretchTypesEnum.esriRasterStretch_HistogramEqualize
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Stretch type: Histogram equalize")
                Case esriRasterStretchTypesEnum.esriRasterStretch_HistogramSpecification
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Stretch type: Histogram specification")
                Case esriRasterStretchTypesEnum.esriRasterStretch_MinimumMaximum
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Stretch type: Minimum maximum")
                Case esriRasterStretchTypesEnum.esriRasterStretch_DefaultFromSource
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Stretch type: Default")
                Case Else
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Stretch type: Other ***** TODO *****")  ' TODO fill in others
            End Select
            If pRasterStretch3.UseGamma Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Gamma: " & pRasterStretch3.GammaValue)
        ElseIf TypeOf pRR Is IRasterUniqueValueRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Raster Unique Value Renderer")
            mxdProps.bRasterUnique = True
            pUVR = pRR
            Dim iHeading As Integer
            Dim iClass As Integer
            If pUVR.ColorScheme <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Colour Scheme: " & pUVR.ColorScheme)
            If pUVR.UseDefaultSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Use Default Symbol")
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Label: " & pUVR.DefaultLabel)
                GetSymbolProps(pUVR.DefaultSymbol, lTabLevel + 1, bSymbolLevels)
            End If
            sw.WriteLine(InsertTabs(lTabLevel) & "Number of headings: " & pUVR.HeadingCount)
            For iHeading = 0 To pUVR.HeadingCount - 1
                sw.WriteLine(InsertTabs(lTabLevel) & "Heading " & iHeading + 1 & "(" & pUVR.Heading(iHeading) & "):")
                sw.WriteLine(InsertTabs(lTabLevel) & "Number of classes: " & pUVR.ClassCount(iHeading))
                For iClass = 0 To pUVR.ClassCount(iHeading) - 1
                    GetSymbolProps(pUVR.Symbol(iHeading, iClass), lTabLevel + 1, bSymbolLevels)
                    ' TODO value objects
                    'sw.WriteLine(InsertTabs(lTabLevel) & "Number of values: " & pUVR.ValueCount(iHeading, iClass))
                    'For i = 0 To pUVR.ValueCount(iHeading, iClass) - 1
                    '    sw.WriteLine(InsertTabs(lTabLevel) & "Value " & i + 1 & "/" & pUVR.ValueCount & _
                    '                 ": " & pUVR.Value(i) & " Label: " & pUVR.Label(pUVR.Value(i)))
                    '    'VB.NET won't let you pass the symbol in if it is nothing
                    '    If pUVR.Symbol(pUVR.Value(i)) Is Nothing Then
                    '        GetSymbolProps(Nothing, lTabLevel + 1, bSymbolLevels)
                    '    Else
                    '        GetSymbolProps(pUVR.Symbol(pUVR.Value(i)), lTabLevel + 1, bSymbolLevels)
                    '    End If
                    'Next
                Next
            Next
        Else
            sw.WriteLine(InsertTabs(lTabLevel) & "Raster Renderer ************ TODO ************")
        End If
    End Sub

    Sub GetRendererProps(ByRef pFR As IFeatureRenderer, ByRef lTabLevel As Integer, ByRef bSymbolLevels As Boolean)
        sw.Flush()
        Dim sTmp As String
        Dim i As Integer
        Dim pSimpRend As ISimpleRenderer
        Dim pBVR As IBivariateRenderer
        Dim pChartRend As IChartRenderer
        Dim pClassBreaksRend As IClassBreaksRenderer
        Dim pRepsRend As IRepresentationRenderer
        Dim pUVR As IUniqueValueRenderer
        Dim pPieChartRend As IPieChartRenderer
        Dim pRendFields As IRendererFields
        Dim pSymArray As ISymbolArray
        If TypeOf pFR Is ISimpleRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Simple Renderer")
            pSimpRend = pFR
            If pSimpRend.Description <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Description: " & pSimpRend.Description)
            If pSimpRend.Label <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Label: " & pSimpRend.Label)
            GetSymbolProps(pSimpRend.Symbol, lTabLevel + 1, bSymbolLevels)
        ElseIf TypeOf pFR Is IBivariateRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Bivariate Renderer")
            pBVR = pFR
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Main Renderer:")
            If Not pBVR.MainRenderer Is Nothing Then GetRendererProps(pBVR.MainRenderer, lTabLevel + 1, bSymbolLevels)
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Variation Renderer:")
            If Not pBVR.VariationRenderer Is Nothing Then GetRendererProps(pBVR.VariationRenderer, lTabLevel + 1, bSymbolLevels)
        ElseIf TypeOf pFR Is IChartRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Chart Renderer")
            pChartRend = pFR
            If pChartRend.ColorScheme <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Colour Scheme: " & pChartRend.ColorScheme)
            If pChartRend.Label <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Label: " & pChartRend.Label)
            If pChartRend.UseOverposter Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Prevent chart overlap")
            Else
                mxdProps.bChartOverlap = True
            End If
            pPieChartRend = pChartRend
            If Not pPieChartRend Is Nothing Then
                If pPieChartRend.ProportionalBySum Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Proportional by sum")
                ElseIf pPieChartRend.ProportionalField <> "" Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Proportional field: " & pPieChartRend.ProportionalField)
                    If pPieChartRend.ProportionalFieldAlias <> "" Then _
                        If Not pPieChartRend.ProportionalFieldAlias.Equals(pPieChartRend.ProportionalField, StringComparison.CurrentCulture) Then _
                            sw.WriteLine(InsertTabs(lTabLevel) & "Proportional field alias: " & pPieChartRend.ProportionalFieldAlias)
                    sw.WriteLine(InsertTabs(lTabLevel) & "Min Size: " & pPieChartRend.MinSize)
                    sw.WriteLine(InsertTabs(lTabLevel) & "Min Value: " & pPieChartRend.MinValue)
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Fixed size")
                    mxdProps.bFixedSize = True
                End If
                If pPieChartRend.FlanneryCompensation Then sw.WriteLine(InsertTabs(lTabLevel) & "Flannery compensation")
            End If
            'put all the field names in an array...
            Dim pFieldNames(ARRAY_SIZE) As String
            pRendFields = pChartRend
            If Not pRendFields Is Nothing Then
                For i = 0 To pRendFields.FieldCount - 1
                    If i < ARRAY_SIZE Then
                        If pRendFields.FieldAlias(i) <> "" Then
                            pFieldNames(i) = InsertTabs(lTabLevel + 1) & "Name: " & pRendFields.Field(i)
                        Else
                            pFieldNames(i) = InsertTabs(lTabLevel + 1) & "Name: " & pRendFields.Field(i) & " Alias: " & pRendFields.FieldAlias(i)
                        End If
                    End If
                Next
            End If
            '... then write the names as we get the symbols
            pSymArray = pChartRend.ChartSymbol
            If Not pSymArray Is Nothing Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Field symbols:")
                For i = 0 To pSymArray.SymbolCount - 1
                    If i < ARRAY_SIZE Then sw.WriteLine(pFieldNames(i))
                    GetSymbolProps(pSymArray.Symbol(i), lTabLevel + 1, bSymbolLevels)
                Next
            End If
            Dim pDataExclusion As IDataExclusion = pChartRend
            If Not pDataExclusion Is Nothing Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Exclusion query: " & pDataExclusion.ExclusionClause)
                If pDataExclusion.ShowExclusionClass Then
                    If pDataExclusion.ExclusionLabel <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Label: " & pDataExclusion.ExclusionLabel)
                    If pDataExclusion.ExclusionDescription <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Description: " & pDataExclusion.ExclusionDescription)
                    GetSymbolProps(pDataExclusion.ExclusionSymbol, lTabLevel, bSymbolLevels)
                End If
            End If
            If Not pChartRend.BaseSymbol Is Nothing Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Base Symbol:")
                GetSymbolProps(pChartRend.BaseSymbol, lTabLevel + 1, bSymbolLevels)
            End If
            'TODO where does leader hook in? here or bar/stacked/piechartsymbol?
            'Dim pMarkerBkgSupport As IMarkerBackgroundSupport = pChartRend.ChartSymbol
            'If Not pMarkerBkgSupport.Background Is Nothing Then
            '    Dim pMarkBkg As IMarkerBackground = pMarkerBkgSupport.Background
            '    If Not pMarkBkg Is Nothing Then
            '        GetSymbolProps(pMarkBkg.MarkerSymbol, lTabLevel, bSymbolLevels)
            '    End If
            'End If
            sw.WriteLine(InsertTabs(lTabLevel) & "Chart Symbol:")
            GetSymbolProps(pChartRend.ChartSymbol, lTabLevel + 1, bSymbolLevels)
        ElseIf TypeOf pFR Is IClassBreaksRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Class Breaks Renderer")
            pClassBreaksRend = pFR
            If Not pClassBreaksRend.BackgroundSymbol Is Nothing Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Background Symbol:")
                GetSymbolProps(pClassBreaksRend.BackgroundSymbol, lTabLevel + 1, bSymbolLevels)
            End If
            If pClassBreaksRend.SortClassesAscending Then sw.WriteLine(InsertTabs(lTabLevel) & "Sort classes ascending")
            If pClassBreaksRend.Field <> vbNullString Then sw.WriteLine(InsertTabs(lTabLevel) & "Field: " & pClassBreaksRend.Field)
            If pClassBreaksRend.NormField <> vbNullString Then sw.WriteLine(InsertTabs(lTabLevel) & "Norm field: " & pClassBreaksRend.NormField)
            sw.WriteLine(InsertTabs(lTabLevel) & "Minimum break: " & pClassBreaksRend.MinimumBreak)
            For i = 0 To pClassBreaksRend.BreakCount - 1
                sw.WriteLine(InsertTabs(lTabLevel) & "Break " & i & ": ")
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Upper bound: " & pClassBreaksRend.Break(i))
                If pClassBreaksRend.Description(i) <> vbNullString Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Description: " & pClassBreaksRend.Description(i))
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Label: " & pClassBreaksRend.Label(i))
                GetSymbolProps(pClassBreaksRend.Symbol(i), lTabLevel + 2, bSymbolLevels)
            Next
        ElseIf TypeOf pFR Is IDotDensityRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Dot Density Renderer")
            sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            'GetSymbolProps pSimpRend.Symbol, lTabLevel + 1, bsymbollevels
        ElseIf TypeOf pFR Is IProportionalSymbolRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Proportional Symbol Renderer")
            sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            'GetSymbolProps pSimpRend.Symbol, lTabLevel + 1, bsymbollevels
        ElseIf TypeOf pFR Is IRepresentationRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Representation Renderer")
            pRepsRend = pFR
            If pRepsRend.DrawInvalidRule Then sw.WriteLine(InsertTabs(lTabLevel) & "Draw invalid rule")
            If pRepsRend.DrawInvisible Then sw.WriteLine(InsertTabs(lTabLevel) & "Draw invisible")
            sw.WriteLine(InsertTabs(lTabLevel) & "Invalid rule color (RGB): " & GetRGB((pRepsRend.InvalidRuleColor)))
            sw.WriteLine(InsertTabs(lTabLevel) & "Invalid rule color (CMYK): " & GetCMYK((pRepsRend.InvalidRuleColor)))
            GetRepresentationClass(pRepsRend.RepresentationClass, lTabLevel + 1)
        ElseIf TypeOf pFR Is IScaleDependentRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Scale Dependent Renderer")
            sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            'GetSymbolProps pSimpRend.Symbol, lTabLevel + 1, bsymbollevels
        ElseIf TypeOf pFR Is IUniqueValueRenderer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Unique Value Renderer")
            pUVR = pFR
            If pUVR.FieldCount = 1 Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Field = " & pUVR.Field(0))
                If IsQualifiedName(pUVR.Field(0)) Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Qualified name!")
            Else
                Dim bQual As Boolean = False
                sTmp = pUVR.Field(0)
                If IsQualifiedName(pUVR.Field(0)) Then bQual = True
                For i = 1 To pUVR.FieldCount - 1
                    sTmp = sTmp & pUVR.FieldDelimiter & " " & pUVR.Field(i)
                    If IsQualifiedName(pUVR.Field(i)) Then bQual = True
                Next
                sw.WriteLine(InsertTabs(lTabLevel) & "Fields = " & sTmp)
                If bQual Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Qualified name!")
            End If
            If pUVR.ColorScheme <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Colour Scheme: " & pUVR.ColorScheme)
            If pUVR.LookupStyleset <> "" Then sw.WriteLine(InsertTabs(lTabLevel) & "Lookup Style: " & pUVR.LookupStyleset)
            If pUVR.UseDefaultSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Use Default Symbol")
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Label: " & pUVR.DefaultLabel)
                GetSymbolProps(pUVR.DefaultSymbol, lTabLevel + 1, bSymbolLevels)
            End If
            sw.WriteLine(InsertTabs(lTabLevel) & "Number of values: " & pUVR.ValueCount)
            For i = 0 To pUVR.ValueCount - 1
                sw.WriteLine(InsertTabs(lTabLevel) & "Value " & i + 1 & "/" & pUVR.ValueCount & _
                             ": " & pUVR.Value(i) & " Label: " & pUVR.Label(pUVR.Value(i)))
                'VB.NET won't let you pass the symbol in if it is nothing
                If pUVR.Symbol(pUVR.Value(i)) Is Nothing Then
                    GetSymbolProps(Nothing, lTabLevel + 1, bSymbolLevels)
                Else
                    GetSymbolProps(pUVR.Symbol(pUVR.Value(i)), lTabLevel + 1, bSymbolLevels)
                End If
            Next
        End If
        Dim pRotRend As IRotationRenderer
        If TypeOf pFR Is IRotationRenderer Then
            pRotRend = pFR
            If pRotRend.RotationField <> "" Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Rotation Field: " & pRotRend.RotationField)
                Select Case pRotRend.RotationType
                    Case esriSymbolRotationType.esriRotateSymbolGeographic
                        sw.WriteLine(InsertTabs(lTabLevel + 1) & "Geographic rotation")
                    Case esriSymbolRotationType.esriRotateSymbolArithmetic
                        sw.WriteLine(InsertTabs(lTabLevel + 1) & "Arithmetic rotation")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel + 1) & "Unknown rotation")
                End Select
            End If
        End If
        '  Dim pTransRend As ITransparencyRenderer
        '  If TypeOf pFR Is ITransparencyRenderer Then
        '    pTransRend = pFR
        '    If pTransRend.TransparencyField <> "" Then sw.writeline(InsertTabs(lTabLevel + 1) & "Transparency Field: " & pTransRend.TransparencyField)
        '  End If

    End Sub

    Sub GetColorRampProps(ByRef pRamp As IColorRamp, ByRef lTabLevel As Integer, ByRef bSymbolLevels As Boolean)
        Dim i As Integer
        Dim pAlgCR As IAlgorithmicColorRamp
        Dim pMultipartCR As IMultiPartColorRamp
        Dim pPresetCR As IPresetColorRamp
        Dim pRandomCR As IRandomColorRamp
        If Not pRamp Is Nothing Then
            mxdProps.bColorRamp = True
            sw.WriteLine(InsertTabs(lTabLevel) & "Color Ramp:")
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Name: " & pRamp.Name)
            If pRamp.Size And Not TypeOf pRamp Is IAlgorithmicColorRamp Then
                ' do not list all colors for algorithmic ramp, just from and to
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Size: " & pRamp.Size)
                For i = 0 To pRamp.Size - 1
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Color " & i & " (RGB): " & GetRGB(pRamp.Color(i)))
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Color " & i & " (CMYK): " & GetCMYK(pRamp.Color(i)))
                Next
            End If
            If TypeOf pRamp Is IAlgorithmicColorRamp Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Algorithmic Color Ramp")
                pAlgCR = pRamp
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Algorithm: " & GetColorRampAlgorithm(pAlgCR.Algorithm))
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "From color (RGB): " & GetRGB(pAlgCR.FromColor))
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "To color (CMYK): " & GetCMYK(pAlgCR.ToColor))
            ElseIf TypeOf pRamp Is IMultiPartColorRamp Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Multipart Color Ramp")
                pMultipartCR = pRamp
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Number of ramps: " & pMultipartCR.NumberOfRamps)
                For i = 0 To pMultipartCR.NumberOfRamps - 1
                    Dim tempRamp As IColorRamp
                    tempRamp = pMultipartCR.Ramp(i)  ' get ramp before passing it in or it starts adding new ramps
                    GetColorRampProps(tempRamp, lTabLevel + 1, bSymbolLevels)
                Next
            ElseIf TypeOf pRamp Is IPresetColorRamp Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Preset Color Ramp")
                pPresetCR = pRamp
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Number of preset colors: " & pPresetCR.NumberOfPresetColors)
                For i = 0 To pPresetCR.NumberOfPresetColors - 1
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Preset color " & i & " (RGB): " & GetRGB(pPresetCR.PresetColor(i)))
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Preset color " & i & " (CMYK): " & GetCMYK(pPresetCR.PresetColor(i)))
                Next
            ElseIf TypeOf pRamp Is IRandomColorRamp Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Random Color Ramp")
                pRandomCR = pRamp
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Start Hue: " & pRandomCR.StartHue)
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "End Hue: " & pRandomCR.EndHue)
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Min Saturation: " & pRandomCR.MinSaturation)
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Max Saturation: " & pRandomCR.MaxSaturation)
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Min Value: " & pRandomCR.MinValue)
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Max Value: " & pRandomCR.MaxValue)
                If pRandomCR.UseSeed Then
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Seed: " & pRandomCR.Seed)
                End If
            End If
        End If
    End Sub

    Sub GetSymbolProps(ByRef pSym As ESRI.ArcGIS.Display.ISymbol, ByRef lTabLevel As Integer, ByRef bSymbolLevels As Boolean)
        sw.Flush()

        If pSym Is Nothing Then
            sw.WriteLine(InsertTabs(lTabLevel) & "No Symbol")
            Return
        End If

        Dim pMapLevel As ESRI.ArcGIS.Display.IMapLevel
        If bSymbolLevels Then
            pMapLevel = pSym
            sw.WriteLine(InsertTabs(lTabLevel) & "Symbol level: " & pMapLevel.MapLevel)
        End If

        Dim i As Integer
        '
        'Marker Symbol
        Dim pMarkSym As ESRI.ArcGIS.Display.IMarkerSymbol
        Dim pSimpMarkSym As ESRI.ArcGIS.Display.ISimpleMarkerSymbol
        Dim pMLyrMarkSym As ESRI.ArcGIS.Display.IMultiLayerMarkerSymbol
        Dim pArrowMarkSym As ESRI.ArcGIS.Display.IArrowMarkerSymbol
        Dim pCharMarkSym As ESRI.ArcGIS.Display.ICharacterMarkerSymbol
        Dim pPicMarkSym As ESRI.ArcGIS.Display.IPictureMarkerSymbol
        Dim pPicDisp As stdole.IPictureDisp = New stdole.StdPictureClass()
        Dim pFontDisp As stdole.IFontDisp = New stdole.StdFontClass()
        Dim pFillSym As ESRI.ArcGIS.Display.IFillSymbol
        Dim pSimpFillSym As ESRI.ArcGIS.Display.ISimpleFillSymbol
        Dim pMLyrFillSym As ESRI.ArcGIS.Display.IMultiLayerFillSymbol
        Dim pGradFillSym As ESRI.ArcGIS.Display.IGradientFillSymbol
        Dim pLineFillSym As ESRI.ArcGIS.Display.ILineFillSymbol
        Dim pMarkFillSym As ESRI.ArcGIS.Display.IMarkerFillSymbol
        Dim pPicFillSym As ESRI.ArcGIS.Display.IPictureFillSymbol
        '        Dim pTexFillSym As ESRI.ArcGIS.Analyst3D.ITextureFillSymbol
        Dim pLineSym As ESRI.ArcGIS.Display.ILineSymbol
        Dim pSimpLineSym As ESRI.ArcGIS.Display.ISimpleLineSymbol
        Dim pMLyrLineSym As ESRI.ArcGIS.Display.IMultiLayerLineSymbol
        Dim pCartoLineSym As ESRI.ArcGIS.Display.ICartographicLineSymbol
        Dim pHashLineSym As IHashLineSymbol
        Dim pPieChartSym As IPieChartSymbol
        Dim p3DChartSym As I3DChartSymbol
        Dim pBarChartSym As IBarChartSymbol
        Dim pStackedChartSym As IStackedChartSymbol
        If TypeOf pSym Is ESRI.ArcGIS.Display.IMarkerSymbol Then
            pMarkSym = pSym
            'Print #InsertTabs(lTabLevel) & "Marker Symbol"
            If TypeOf pMarkSym Is ESRI.ArcGIS.Display.ISimpleMarkerSymbol Then
                pSimpMarkSym = pMarkSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Simple Marker Symbol")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                If pSimpMarkSym.Outline Then
                    If pSimpMarkSym.Color Is Nothing Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "No outline color")
                    Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Outline color (RGB): " & GetRGB(pSimpMarkSym.Color))
                        sw.WriteLine(InsertTabs(lTabLevel) & "Outline color (CMYK): " & GetCMYK(pSimpMarkSym.Color))
                    End If
                    sw.WriteLine(InsertTabs(lTabLevel) & "Outline size: " & pSimpMarkSym.OutlineSize)
                End If
                sw.WriteLine(InsertTabs(lTabLevel) & "Style: " & pSimpMarkSym.Style)
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Display.IMultiLayerMarkerSymbol Then
                pMLyrMarkSym = pMarkSym
                sw.WriteLine(InsertTabs(lTabLevel) & "MultiLayer Marker Symbol")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                For i = 0 To pMLyrMarkSym.LayerCount - 1
                    sw.WriteLine(InsertTabs(lTabLevel) & "Symbol " & i + 1 & "/" & pMLyrMarkSym.LayerCount)
                    GetSymbolProps(pMLyrMarkSym.Layer(i), lTabLevel + 1, bSymbolLevels)
                Next
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Display.IArrowMarkerSymbol Then
                pArrowMarkSym = pMarkSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Arrow Marker Symbol")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                If Math.Abs(pArrowMarkSym.Length) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Length: " & pArrowMarkSym.Length)
                If Math.Abs(pArrowMarkSym.Width) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Width: " & pArrowMarkSym.Width)
                Select Case pArrowMarkSym.Style
                    Case ESRI.ArcGIS.Display.esriArrowMarkerStyle.esriAMSPlain
                        sw.WriteLine(InsertTabs(lTabLevel) & "AMS Plain style")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Unknown AMS style")
                End Select
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Analyst3D.ICharacterMarker3DSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "3D Character Marker Symbol")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Display.ICharacterMarkerSymbol Then
                pCharMarkSym = pSym
                pFontDisp = pCharMarkSym.Font
                sw.WriteLine(InsertTabs(lTabLevel) & "Character Marker Symbol")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "Font: " & pFontDisp.Name)
                sw.WriteLine(InsertTabs(lTabLevel) & "Font Display Size: " & pFontDisp.Size) '.SizeInPoints)
                sw.WriteLine(InsertTabs(lTabLevel) & GetCharSet((pFontDisp.Charset))) '.GdiCharSet())))
                If pFontDisp.Bold Then sw.WriteLine(InsertTabs(lTabLevel) & "Bold")
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Bold weight " & pFontDisp.Weight)
                If pFontDisp.Italic Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Italic")
                If pFontDisp.Underline Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Underline")
                If pFontDisp.Strikethrough Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Strikethrough")
                sw.WriteLine(InsertTabs(lTabLevel) & "Index: " & pCharMarkSym.CharacterIndex)
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Analyst3D.IMarker3DSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "3D Marker Symbol")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Display.IPictureMarkerSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Picture Marker Symbol")
                pPicMarkSym = pSym
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "Background color (RGB): " & GetRGB(pPicMarkSym.BackgroundColor))
                sw.WriteLine(InsertTabs(lTabLevel) & "Background color (CMYK): " & GetCMYK(pPicMarkSym.BackgroundColor))
                sw.WriteLine(InsertTabs(lTabLevel) & "Transparency color (RGB): " & GetRGB(pPicMarkSym.BitmapTransparencyColor))
                sw.WriteLine(InsertTabs(lTabLevel) & "Transparency color (CMYK): " & GetCMYK(pPicMarkSym.BitmapTransparencyColor))
                If pPicMarkSym.SwapForeGroundBackGroundColor Then sw.WriteLine(InsertTabs(lTabLevel) & "Swap foreground and background color")
                pPicDisp = pPicMarkSym.Picture
                sw.WriteLine(InsertTabs(lTabLevel) & "Size: " & pPicDisp.Height & " x " & pPicDisp.Width)
                sw.WriteLine(InsertTabs(lTabLevel) & "Type: " & GetPicType(pPicDisp.Type))
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Display.IBarChartSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Bar Chart Symbol")
                mxdProps.bBarChart = True
                GetMarkerSymbolProps(pMarkSym, lTabLevel, False, False)
                pBarChartSym = pMarkSym
                If pBarChartSym.VerticalBars Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Column orientation")
                    mxdProps.bColumnOrient = True
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Bar orientation")
                    mxdProps.bBarOrient = True
                End If
                sw.WriteLine(InsertTabs(lTabLevel) & "Width: " & pBarChartSym.Width)
                sw.WriteLine(InsertTabs(lTabLevel) & "Spacing: " & pBarChartSym.Spacing)
                If pBarChartSym.ShowAxes Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Axes:")
                    GetSymbolProps(pBarChartSym.Axes, lTabLevel + 1, bSymbolLevels)
                End If
                p3DChartSym = pMarkSym
                If p3DChartSym.Display3D Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Display 3D")
                    mxdProps.b3DChart = True
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Thickness: " & p3DChartSym.Thickness)
                End If
                'Dim pMarkerBkgSupport As IMarkerBackgroundSupport = pBarChartSym
                'GetSymbolProps(pMarkerBkgSupport.Background.MarkerSymbol, lTabLevel, bSymbolLevels)
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Display.IStackedChartSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Stacked Chart Symbol")
                mxdProps.bStackedChart = True
                GetMarkerSymbolProps(pMarkSym, lTabLevel, False, False)
                pStackedChartSym = pMarkSym
                If pStackedChartSym.VerticalBar Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Column orientation")
                    mxdProps.bColumnOrient = True
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Bar orientation")
                    mxdProps.bBarOrient = True
                End If
                sw.WriteLine(InsertTabs(lTabLevel) & "Width: " & pStackedChartSym.Width)
                If pStackedChartSym.Fixed Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Fixed length")
                    mxdProps.bFixedSize = True
                End If
                If pStackedChartSym.UseOutline Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Outline:")
                    GetSymbolProps(pStackedChartSym.Outline, lTabLevel + 1, bSymbolLevels)
                End If
                p3DChartSym = pMarkSym
                If p3DChartSym.Display3D Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Display 3D")
                    mxdProps.b3DChart = True
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Thickness: " & p3DChartSym.Thickness)
                End If
                'Dim pMarkerBkgSupport As IMarkerBackgroundSupport = pStackedChartSym
                'GetSymbolProps(pMarkerBkgSupport.Background.MarkerSymbol, lTabLevel, bSymbolLevels)
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Display.IPieChartSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Pie Chart Symbol")
                mxdProps.bPieChart = True
                GetMarkerSymbolProps(pMarkSym, lTabLevel, False)
                pPieChartSym = pMarkSym
                If pPieChartSym.Clockwise Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Geographic orientation")
                    mxdProps.bGeogOrient = True
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Arithmetic orientation")
                    mxdProps.bArithOrient = True
                End If
                If pPieChartSym.UseOutline Then
                    GetLineSymbolProps(pPieChartSym.Outline, lTabLevel)
                End If
                p3DChartSym = pMarkSym
                If p3DChartSym.Display3D Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Display 3D")
                    mxdProps.b3DChart = True
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Thickness: " & p3DChartSym.Thickness)
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Tilt: " & p3DChartSym.Tilt)
                End If
                'Dim pMarkerBkgSupport As IMarkerBackgroundSupport = pPieChartSym
                'If Not pMarkerBkgSupport.Background Is Nothing Then _
                '    GetSymbolProps(pMarkerBkgSupport.Background.MarkerSymbol, lTabLevel, bSymbolLevels)
            ElseIf TypeOf pMarkSym Is ESRI.ArcGIS.Analyst3D.ISimpleMarker3DSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "3D Simple Marker Symbol")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            Else
                sw.WriteLine("***** Other marker symbol *****")
                GetMarkerSymbolProps(pMarkSym, lTabLevel)
            End If
            '
            'Fill Symbol
        ElseIf TypeOf pSym Is ESRI.ArcGIS.Display.IFillSymbol Then
            pFillSym = pSym
            'Print #InsertTabs(lTabLevel) & "Fill Symbol"
            If TypeOf pFillSym Is ESRI.ArcGIS.Display.ISimpleFillSymbol Then
                pSimpFillSym = pFillSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Simple Fill Symbol")
                mxdProps.bSimpleFill = True
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
                Select Case pSimpFillSym.Style
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSSolid
                        sw.WriteLine(InsertTabs(lTabLevel) & "Solid style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSNull
                        sw.WriteLine(InsertTabs(lTabLevel) & "No style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSHollow
                        sw.WriteLine(InsertTabs(lTabLevel) & "Hollow style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSHorizontal
                        sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSVertical
                        sw.WriteLine(InsertTabs(lTabLevel) & "Vertical style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSForwardDiagonal
                        sw.WriteLine(InsertTabs(lTabLevel) & "Forward Diagonal style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSBackwardDiagonal
                        sw.WriteLine(InsertTabs(lTabLevel) & "Backward Diagonal style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSCross
                        sw.WriteLine(InsertTabs(lTabLevel) & "Cross style")
                    Case ESRI.ArcGIS.Display.esriSimpleFillStyle.esriSFSDiagonalCross
                        sw.WriteLine(InsertTabs(lTabLevel) & "Diagonal Cross style")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Unknown style")
                End Select
            ElseIf TypeOf pFillSym Is ESRI.ArcGIS.Display.IMultiLayerFillSymbol Then
                pMLyrFillSym = pFillSym
                sw.WriteLine(InsertTabs(lTabLevel) & "MultiLayer Fill Symbol")
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
                For i = 0 To pMLyrFillSym.LayerCount - 1
                    sw.WriteLine(InsertTabs(lTabLevel) & "Symbol " & i + 1 & "/" & pMLyrFillSym.LayerCount)
                    GetSymbolProps(pMLyrFillSym.Layer(i), lTabLevel + 1, bSymbolLevels)
                Next
            ElseIf TypeOf pFillSym Is ESRI.ArcGIS.Display.IGradientFillSymbol Then
                pGradFillSym = pFillSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Gradient Fill Symbol")
                mxdProps.bGradientFill = True
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
                If pGradFillSym.IntervalCount <> 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Interval Count: " & pGradFillSym.IntervalCount)
                GetColorRampProps(pGradFillSym.ColorRamp, lTabLevel, bSymbolLevels)
                If Math.Abs(pGradFillSym.GradientAngle) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Gradient Angle: " & pGradFillSym.GradientAngle)
                If Math.Abs(pGradFillSym.GradientPercentage) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Gradient Percentage: " & pGradFillSym.GradientPercentage)
                Select Case pGradFillSym.Style
                    Case ESRI.ArcGIS.Display.esriGradientFillStyle.esriGFSBuffered
                        sw.WriteLine(InsertTabs(lTabLevel) & "GFS Buffered style")
                    Case ESRI.ArcGIS.Display.esriGradientFillStyle.esriGFSCircular
                        sw.WriteLine(InsertTabs(lTabLevel) & "GFS Circular style")
                    Case ESRI.ArcGIS.Display.esriGradientFillStyle.esriGFSLinear
                        sw.WriteLine(InsertTabs(lTabLevel) & "GFS Linear style")
                    Case ESRI.ArcGIS.Display.esriGradientFillStyle.esriGFSRectangular
                        sw.WriteLine(InsertTabs(lTabLevel) & "GFS Rectangular style")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Unknown style")
                End Select
            ElseIf TypeOf pFillSym Is ESRI.ArcGIS.Display.ILineFillSymbol Then
                pLineFillSym = pFillSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Line Fill Symbol")
                mxdProps.bLineFill = True
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
                If Math.Abs(pLineFillSym.Angle) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Angle: " & pLineFillSym.Angle)
                If Math.Abs(pLineFillSym.Offset) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Offset: " & pLineFillSym.Offset)
                If Math.Abs(pLineFillSym.Separation) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Separation: " & pLineFillSym.Separation)
                GetSymbolProps(pLineFillSym.LineSymbol, lTabLevel + 1, bSymbolLevels)
            ElseIf TypeOf pFillSym Is ESRI.ArcGIS.Display.IMarkerFillSymbol Then
                pMarkFillSym = pFillSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Marker Fill Symbol")
                mxdProps.bMarkerFill = True
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
                Select Case pMarkFillSym.Style
                    Case ESRI.ArcGIS.Display.esriMarkerFillStyle.esriMFSGrid
                        sw.WriteLine(InsertTabs(lTabLevel) & "Grid placement")
                    Case ESRI.ArcGIS.Display.esriMarkerFillStyle.esriMFSRandom
                        sw.WriteLine(InsertTabs(lTabLevel) & "Random placement")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Unknown placement")
                End Select
                GetSymbolProps(pMarkFillSym.MarkerSymbol, lTabLevel + 1, bSymbolLevels)
            ElseIf TypeOf pFillSym Is ESRI.ArcGIS.Display.IPictureFillSymbol Then
                pPicFillSym = pFillSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Picture Fill Symbol")
                mxdProps.bPictureFill = True
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
                If Math.Abs(pPicFillSym.Angle) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Angle: " & pPicFillSym.Angle)
                If Math.Abs(pPicFillSym.XScale) - 1.0 > 0 Or Math.Abs(pPicFillSym.YScale) - 1.0 > 0 Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "X Scale: " & pPicFillSym.XScale)
                    sw.WriteLine(InsertTabs(lTabLevel) & "Y Scale: " & pPicFillSym.YScale)
                End If
                If Not pPicFillSym.BackgroundColor Is Nothing Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Background Color (RGB): " & GetRGB(pPicFillSym.BackgroundColor))
                    sw.WriteLine(InsertTabs(lTabLevel) & "Background Color (CMYK): " & GetCMYK(pPicFillSym.BackgroundColor))
                End If
                If Not pPicFillSym.BitmapTransparencyColor Is Nothing Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Bitmap Transparency Color (RGB): " & GetRGB(pPicFillSym.BitmapTransparencyColor))
                    sw.WriteLine(InsertTabs(lTabLevel) & "Bitmap Transparency Color (CMYK): " & GetCMYK(pPicFillSym.BitmapTransparencyColor))
                End If
                If pPicFillSym.SwapForeGroundBackGroundColor Then sw.WriteLine(InsertTabs(lTabLevel) & "Swap foreground and background color")
                'If Not pPicFillSym.Picture Is Nothing Then
                '  Dim pPicDisp As IPictureDisp
                '  Set pPicDisp = pPicFillSym.Picture
                'now what?
                'End If
            ElseIf TypeOf pFillSym Is ESRI.ArcGIS.Analyst3D.ITextureFillSymbol Then
                '                pTexFillSym = pFillSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Texture Fill Symbol")
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
                sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            ElseIf TypeOf pFillSym Is IColorSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Color Symbol")
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
            Else
                sw.WriteLine("***** Other fill symbol *****")
                GetFillSymbolProps(pFillSym, lTabLevel, bSymbolLevels)
            End If
            '
            'Line Symbol
        ElseIf TypeOf pSym Is ESRI.ArcGIS.Display.ILineSymbol Then
            pLineSym = pSym
            If TypeOf pLineSym Is ESRI.ArcGIS.Display.IHashLineSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Hash Line Symbol")
                pHashLineSym = pLineSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Angle: " & pHashLineSym.Angle)
                'don't try this, you'll end up in an infinite loop!
                'GetLineSymbolProps(pHashLineSym, lTabLevel)
            End If
            If TypeOf pLineSym Is ESRI.ArcGIS.Display.ISimpleLineSymbol Then
                pSimpLineSym = pLineSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Simple Line Symbol")
                GetLineSymbolProps(pLineSym, lTabLevel)
                Select Case pSimpLineSym.Style
                    Case ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSDash
                        sw.WriteLine(InsertTabs(lTabLevel) & "Dash style")
                    Case ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSDashDot
                        sw.WriteLine(InsertTabs(lTabLevel) & "Dash Dot style")
                    Case ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSDashDotDot
                        sw.WriteLine(InsertTabs(lTabLevel) & "Dash Dot Dot style")
                    Case ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSDot
                        sw.WriteLine(InsertTabs(lTabLevel) & "Dot style")
                    Case ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSInsideFrame
                        sw.WriteLine(InsertTabs(lTabLevel) & "Inside Frame style")
                    Case ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSNull
                        sw.WriteLine(InsertTabs(lTabLevel) & "No style")
                    Case ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSSolid
                        sw.WriteLine(InsertTabs(lTabLevel) & "Solid style")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Unknown style")
                End Select
            ElseIf TypeOf pLineSym Is ESRI.ArcGIS.Display.IMultiLayerLineSymbol Then
                pMLyrLineSym = pLineSym
                sw.WriteLine(InsertTabs(lTabLevel) & "MultiLayer Line Symbol")
                GetLineSymbolProps(pLineSym, lTabLevel)
                For i = 0 To pMLyrLineSym.LayerCount - 1
                    sw.WriteLine(InsertTabs(lTabLevel) & "Symbol " & i + 1 & "/" & pMLyrLineSym.LayerCount)
                    GetSymbolProps(pMLyrLineSym.Layer(i), lTabLevel + 1, bSymbolLevels)
                Next
            ElseIf TypeOf pLineSym Is ESRI.ArcGIS.Display.ICartographicLineSymbol Then
                pCartoLineSym = pLineSym
                sw.WriteLine(InsertTabs(lTabLevel) & "Cartographic Line Symbol")
                GetLineSymbolProps(pLineSym, lTabLevel)
                Select Case pCartoLineSym.Cap
                    Case ESRI.ArcGIS.Display.esriLineCapStyle.esriLCSButt
                        sw.WriteLine(InsertTabs(lTabLevel) & "Butt Line Caps")
                    Case ESRI.ArcGIS.Display.esriLineCapStyle.esriLCSRound
                        sw.WriteLine(InsertTabs(lTabLevel) & "Round Line Caps")
                    Case ESRI.ArcGIS.Display.esriLineCapStyle.esriLCSSquare
                        sw.WriteLine(InsertTabs(lTabLevel) & "Square Line Caps")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Unknown caps")
                End Select
                Select Case pCartoLineSym.Join
                    Case ESRI.ArcGIS.Display.esriLineJoinStyle.esriLJSMitre
                        sw.WriteLine(InsertTabs(lTabLevel) & "Mitre Line Joins")
                    Case ESRI.ArcGIS.Display.esriLineJoinStyle.esriLJSRound
                        sw.WriteLine(InsertTabs(lTabLevel) & "Round Line Joins")
                    Case ESRI.ArcGIS.Display.esriLineJoinStyle.esriLJSBevel
                        sw.WriteLine(InsertTabs(lTabLevel) & "Bevel Line Joins")
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Unknown joins")
                End Select
                GetLineProps(pCartoLineSym, lTabLevel)
            ElseIf TypeOf pLineSym Is ESRI.ArcGIS.Display.IPictureLineSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Picture Line Symbol")
                GetLineSymbolProps(pLineSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            ElseIf TypeOf pLineSym Is ESRI.ArcGIS.Analyst3D.ISimpleLine3DSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "3D Simple Line Symbol")
                GetLineSymbolProps(pLineSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            ElseIf TypeOf pLineSym Is ESRI.ArcGIS.Analyst3D.ITextureLineSymbol Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Texture Line Symbol")
                GetLineSymbolProps(pLineSym, lTabLevel)
                sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
            Else
                sw.WriteLine(InsertTabs(lTabLevel) & "***** Other line symbol *****")
                GetLineSymbolProps(pLineSym, lTabLevel)
            End If
        Else
            sw.WriteLine(InsertTabs(lTabLevel) & "***** Other symbol *****")
        End If
    End Sub

    'inherited properties on all types of marker symbol
    'do not try to get color or angle for some chart symbols (set bools to false)
    Sub GetMarkerSymbolProps(ByRef pMarkSym As ESRI.ArcGIS.Display.IMarkerSymbol, ByRef lTabLevel As Integer, _
                             Optional ByVal bColor As Boolean = True, Optional ByVal bAngle As Boolean = True)
        sw.Flush()
        sw.WriteLine(InsertTabs(lTabLevel) & "Size: " & pMarkSym.Size)
        If bColor Then
            If pMarkSym.Color Is Nothing Then
                sw.WriteLine(InsertTabs(lTabLevel) & "No Color")
            Else
                sw.WriteLine(InsertTabs(lTabLevel) & "Color (RGB): " & GetRGB(pMarkSym.Color))
                sw.WriteLine(InsertTabs(lTabLevel) & "Color (CMYK): " & GetCMYK(pMarkSym.Color))
            End If
        End If
        If bAngle Then
            If Math.Abs(pMarkSym.Angle) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Angle: " & pMarkSym.Angle)
        End If
        If Math.Abs(pMarkSym.XOffset) > 0 Or Math.Abs(pMarkSym.YOffset) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Offset: " & pMarkSym.XOffset & ", " & pMarkSym.YOffset)
    End Sub

    'inherited properties on all types of fill symbol
    Sub GetFillSymbolProps(ByRef pFillSym As ESRI.ArcGIS.Display.IFillSymbol, ByRef lTabLevel As Integer, ByRef bSymbolLevels As Boolean)
        sw.Flush()
        If pFillSym.Color Is Nothing Then
            sw.WriteLine(InsertTabs(lTabLevel) & "No Color")
        Else
            sw.WriteLine(InsertTabs(lTabLevel) & "Color (RGB): " & GetRGB(pFillSym.Color))
            sw.WriteLine(InsertTabs(lTabLevel) & "Color (CMYK): " & GetCMYK(pFillSym.Color))
        End If
        If Not TypeOf pFillSym Is IColorSymbol Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Outline: ")
            GetSymbolProps(pFillSym.Outline, lTabLevel + 1, bSymbolLevels)
        End If
    End Sub

    'inherited properties on all types of line symbol
    Sub GetLineSymbolProps(ByRef pLineSym As ESRI.ArcGIS.Display.ILineSymbol, ByRef lTabLevel As Integer)
        sw.Flush()
        If pLineSym.Color Is Nothing Then
            sw.WriteLine(InsertTabs(lTabLevel) & "No Color")
        Else
            sw.WriteLine(InsertTabs(lTabLevel) & "Color (RGB): " & GetRGB(pLineSym.Color))
            sw.WriteLine(InsertTabs(lTabLevel) & "Color (CMYK): " & GetCMYK(pLineSym.Color))
        End If
        sw.WriteLine(InsertTabs(lTabLevel) & "Width: " & pLineSym.Width)
    End Sub

    Sub GetLineProps(ByRef pLineProps As ESRI.ArcGIS.Display.ILineProperties, ByRef lTabLevel As Integer)
        sw.Flush()
        If pLineProps.Offset Or pLineProps.Flip Or pLineProps.DecorationOnTop Or pLineProps.LineStartOffset Or _
            (Not pLineProps.LineDecoration Is Nothing) Then sw.WriteLine(InsertTabs(lTabLevel) & "Line Properties:")
        If pLineProps.Offset Then sw.WriteLine(InsertTabs(lTabLevel) & "Offset: " & pLineProps.Offset)
        If pLineProps.Flip Then sw.WriteLine(InsertTabs(lTabLevel) & "Flip")
        If pLineProps.DecorationOnTop Then sw.WriteLine(InsertTabs(lTabLevel) & "Decoration on top")
        If pLineProps.LineStartOffset Then sw.WriteLine(InsertTabs(lTabLevel) & "Line start offset: " & pLineProps.LineStartOffset)
        Dim pLineDec As ESRI.ArcGIS.Display.ILineDecoration
        Dim pLineDecElem As ESRI.ArcGIS.Display.ILineDecorationElement
        Dim pSimpLineDec As ESRI.ArcGIS.Display.ISimpleLineDecorationElement
        Dim i As Integer
        If Not pLineProps.LineDecoration Is Nothing Then
            pLineDec = pLineProps.LineDecoration
            For i = 0 To pLineDec.ElementCount - 1
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Decoration element " & i + 1 & "/" & pLineDec.ElementCount & ":")
                pLineDecElem = pLineDec.Element(i)
                pSimpLineDec = pLineDecElem
                If pSimpLineDec.FlipAll Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Flip all")
                If pSimpLineDec.FlipFirst Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Flip first")
                If pSimpLineDec.Rotate Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Flip all")
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Symbol:")
                GetSymbolProps(pSimpLineDec.MarkerSymbol, lTabLevel + 2, False)
            Next
        End If
    End Sub

    Sub GetRepresentationClass(ByRef pRepClass As ESRI.ArcGIS.Geodatabase.IRepresentationClass, ByRef lTabLevel As Integer)
        sw.WriteLine(InsertTabs(lTabLevel) & "Representation Class:")
        sw.WriteLine(InsertTabs(lTabLevel) & "************ TODO ************")
    End Sub

    Sub GetAnnoProps(ByRef pLayer As ILayer, ByRef pAnnotateLPColl As IAnnotateLayerPropertiesCollection, _
                       ByRef AnnoPropID As UID, ByRef bCheckLabelEngine As Boolean, ByRef lTabLevel As Integer)
        sw.Flush()
        Dim pGeoFL As IGeoFeatureLayer = Nothing
        Dim pAnnotateLP As IAnnotateLayerProperties = Nothing
        Dim pLabelEngineLayerProperties As ILabelEngineLayerProperties2
        Dim lLabelClass As Integer
        Dim sExpress As String
        Dim sTmp As String

        If TypeOf pLayer Is IGeoFeatureLayer Then pGeoFL = pLayer
        sw.WriteLine(InsertTabs(lTabLevel) & "Number of label classes: " & pAnnotateLPColl.Count)
        For lLabelClass = 0 To pAnnotateLPColl.Count - 1
            pAnnotateLPColl.QueryItem(lLabelClass, pAnnotateLP, Nothing, Nothing)

            'name, scale range and visibility
            sw.WriteLine(vbCrLf & InsertTabs(lTabLevel) & pLayer.Name & " label class " & lLabelClass + 1 & "/" & pAnnotateLPColl.Count)
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Name: " & pAnnotateLP.Class & ", Scale range: " & pAnnotateLP.AnnotationMinimumScale & " - " & pAnnotateLP.AnnotationMaximumScale)
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Display: " & pAnnotateLP.DisplayAnnotation)
            'count label class even if not visible
            If AnnoPropID Is Nothing Then
                'count label classes in anno layers separately
                mxdProps.lAnnoLabelClassCount = mxdProps.lAnnoLabelClassCount + 1
            Else
                mxdProps.lLabelClassCount = mxdProps.lLabelClassCount + 1 'feature layers only
            End If
            'if visible or looking at all layers
            If pAnnotateLP.DisplayAnnotation Or bAllLayers Then
                pLabelEngineLayerProperties = pAnnotateLP
                If Math.Abs(pAnnotateLP.AnnotationMinimumScale) > 0 Or Math.Abs(pAnnotateLP.AnnotationMaximumScale) > 0 Then mxdProps.bScaleRanges = True

                'label class properties
                If pAnnotateLP.WhereClause <> "" Then
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "SQL query: " & pAnnotateLP.WhereClause)
                    If IsQualifiedName(pAnnotateLP.WhereClause) Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Qualified name!")
                    mxdProps.bSQL = True
                End If
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Layer priority: " & pAnnotateLP.Priority)
                If pAnnotateLP.Priority <> mxdProps.lLabelClassCount And pAnnotateLP.Priority <> -1 Then mxdProps.bLabelPriority = True
                If pAnnotateLP.Priority = -1 Then mxdProps.bUninitPriority = True

                'label expression (only show first part if too long, unless show full is true)
                sExpress = pLabelEngineLayerProperties.Expression
                If pLabelEngineLayerProperties.IsExpressionSimple Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Simple expression")
                If Len(sExpress) > 2 And ((InStr(1, sExpress, "[") = 1 And InStr(1, sExpress, "&") <> 0) Or InStr(1, sExpress, "[") <> 1) Then
                    If Len(sExpress) < 250 Then
                        sTmp = Replace(sExpress, vbCrLf, vbCrLf & InsertTabs(lTabLevel + 1) & "                  ")
                    Else
                        If bShowFullExp Then
                            sTmp = vbCrLf & sExpress
                        Else
                            sTmp = Left(sExpress, InStr(sExpress, vbCrLf)) & "..."
                        End If
                    End If
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Label expression: " & sTmp)
                    mxdProps.bLabelExpression = True
                    If InStr(1, sExpress, "</") Then
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Contains tags")
                        mxdProps.bTags = True
                    End If
                    If InStr(1, sExpress, "<BSE>", CompareMethod.Text) Then
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Contains Base tag")
                        mxdProps.bBaseTag = True
                    End If
                    If InStr(1, sExpress, "<BAS>", CompareMethod.Text) Then sw.WriteLine("BAS tag!")
                    If InStr(1, sExpress, "&amp;", CompareMethod.Text) Or InStr(1, sExpress, "&quot;", CompareMethod.Text) _
                        Or InStr(1, sExpress, "&apos;", CompareMethod.Text) Or InStr(1, sExpress, "&lt;", CompareMethod.Text) _
                        Or InStr(1, sExpress, "&gt;", CompareMethod.Text) Then
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Contains HTML entities")
                        mxdProps.bHTMLEnt = True
                    End If
                Else
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Label: " & sExpress)
                End If
                If IsQualifiedName(sExpress) Then sw.WriteLine(InsertTabs(lTabLevel + 2) & "Qualified name!")
                If Not pGeoFL Is Nothing Then
                    ' Check for coded value domain and look up if can easily get field
                    Dim pCodedValueAttr As ICodedValueAttributes = CType(pLabelEngineLayerProperties.ExpressionParser, ICodedValueAttributes)
                    If pCodedValueAttr.UseCodedValue Then
                        ' Get the field(s) from the expression (each string between square brackets)
                        Dim pos As Integer = 0
                        Dim sFieldNames(ARRAY_SIZE) As String
                        Dim iFieldName As Integer = 0
                        While pos >= 0 And pos < sExpress.Length
                            pos = sExpress.IndexOf("[", pos, StringComparison.Ordinal)
                            If pos < 0 Then Exit While
                            If pos < sExpress.Length Then
                                'find close bracket, exit if not found
                                Dim pos2 As Integer = sExpress.IndexOf("]", pos, StringComparison.Ordinal)
                                If pos2 < 0 Then Exit While
                                'add unique names to array
                                sFieldNames(iFieldName) = sExpress.Substring(pos + 1, pos2 - pos - 1)
                                AddIfUnique(iFieldName, sFieldNames, ARRAY_SIZE)
                            End If
                            pos = pos + 1
                        End While
                        'get domains for unique field names only
                        For i As Integer = 0 To iFieldName - 1
                            GetCodedValueDomain(pGeoFL, sFieldNames(i), lTabLevel)
                        Next
                    End If 'use coded value
                End If 'pgeofl is nothing

                'text symbol
                GetTextSymbolProps((pLabelEngineLayerProperties.Symbol), lTabLevel + 2)

                'label props
                If Not AnnoPropID Is Nothing Then
                    If bCheckLabelEngine Then
                        Dim pMaplexUID As UID = New UID
                        Dim pLayerUID As UID = New UID
                        pMaplexUID.Value = "{20664808-0D1C-11D2-A26F-080009B6F22B}"
                        pLayerUID.Value = AnnoPropID
                        If pMaplexUID.Compare(pLayerUID) Then
                            mxdProps.bMapIsMLE = True
                        Else
                            mxdProps.bMapIsSLE = True
                        End If
                    End If

                    If mxdProps.bMapIsMLE Then _
                      GetMLEProps(pLabelEngineLayerProperties, (pLabelEngineLayerProperties.Symbol), lTabLevel + 2)
                    If mxdProps.bMapIsSLE Then _
                      GetSLEProps(pLabelEngineLayerProperties, lTabLevel + 2)
                End If

            End If 'display or show all: label class
        Next lLabelClass


    End Sub

    Sub GetTextSymbolProps(ByRef pTextSym As ITextSymbol, ByRef lTabLevel As Integer)
        sw.Flush()
        Dim pSimpleTextSymbol As ISimpleTextSymbol
        Dim pFormTextSym As IFormattedTextSymbol
        Dim pTextBackground As ITextBackground
        Dim pBalloonCallout As IBalloonCallout
        Dim pLineCallout As ILineCallout
        Dim pSimpleLineCallout As ISimpleLineCallout
        Dim pMarkerTextBkg As IMarkerTextBackground
        Dim pMarkerSymbol As IMarkerSymbol
        Dim pMask As IMask
        Dim pMaskSym As IFillSymbol
        Dim pCharOrientation As ICharacterOrientation
        Dim pFontDisp As stdole.IFontDisp = New stdole.StdFontClass()
        pSimpleTextSymbol = pTextSym
        pFormTextSym = pTextSym
        pTextBackground = pFormTextSym.Background
        pMask = pTextSym
        pCharOrientation = pTextSym
        pFontDisp = pTextSym.Font

        sw.WriteLine(InsertTabs(lTabLevel - 1) & "Text symbol properties:")
        'General tab
        sw.WriteLine(InsertTabs(lTabLevel) & "Font: " & pTextSym.Font.Name)
        sw.WriteLine(InsertTabs(lTabLevel) & "Text Symbol Size: " & pTextSym.Size)
        sw.WriteLine(InsertTabs(lTabLevel) & "Font Display Size: " & pFontDisp.Size) '.SizeInPoints)
        sw.WriteLine(InsertTabs(lTabLevel) & "Color (RGB): " & GetRGB((pTextSym.Color)))
        sw.WriteLine(InsertTabs(lTabLevel) & "Color (CMYK): " & GetCMYK((pTextSym.Color)))
        sw.WriteLine(InsertTabs(lTabLevel) & GetCharSet((pFontDisp.Charset))) '.GdiCharSet())))
        sw.WriteLine(InsertTabs(lTabLevel) & "Angle: " & pTextSym.Angle)
        If pFontDisp.Bold Then sw.WriteLine(InsertTabs(lTabLevel) & "Bold")
        sw.WriteLine(InsertTabs(lTabLevel + 1) & "Bold weight " & pFontDisp.Weight)
        If pFontDisp.Italic Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Italic")
        If pFontDisp.Underline Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Underline")
        If pFontDisp.Strikethrough Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Strikethrough")
        If Math.Abs(pSimpleTextSymbol.XOffset) > 0 Or Math.Abs(pSimpleTextSymbol.YOffset) > 0 Then mxdProps.bXYOffset = True
        If Math.Abs(pSimpleTextSymbol.XOffset) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "X offset: " & pSimpleTextSymbol.XOffset)
        If Math.Abs(pSimpleTextSymbol.YOffset) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Y offset: " & pSimpleTextSymbol.YOffset)
        If pTextSym.HorizontalAlignment = esriTextHorizontalAlignment.esriTHACenter Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal alignment: Center")
        ElseIf pTextSym.HorizontalAlignment = esriTextHorizontalAlignment.esriTHAFull Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal alignment: Full")
        ElseIf pTextSym.HorizontalAlignment = esriTextHorizontalAlignment.esriTHALeft Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal alignment: Left")
        ElseIf pTextSym.HorizontalAlignment = esriTextHorizontalAlignment.esriTHARight Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal alignment: Right")
        End If
        If pTextSym.VerticalAlignment = esriTextVerticalAlignment.esriTVACenter Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Vertical alignment: Center")
        ElseIf pTextSym.VerticalAlignment = esriTextVerticalAlignment.esriTVATop Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Vertical alignment: Top")
        ElseIf pTextSym.VerticalAlignment = esriTextVerticalAlignment.esriTVABaseline Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Vertical alignment: Baseline")
        ElseIf pTextSym.VerticalAlignment = esriTextVerticalAlignment.esriTVABottom Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Vertical alignment: Bottom")
        End If
        If pSimpleTextSymbol.RightToLeft() Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Right to Left")
            mxdProps.bRighttoLeft = True
        End If
        'Formatted Text tab
        If pFormTextSym.Position <> esriTextPosition.esriTPNormal Then
            mxdProps.bTextPosition = True
            If pFormTextSym.Position = esriTextPosition.esriTPSubscript Then sw.WriteLine(InsertTabs(lTabLevel) & "Subscript")
            If pFormTextSym.Position = esriTextPosition.esriTPSuperscript Then sw.WriteLine(InsertTabs(lTabLevel) & "Superscript")
        End If
        If pFormTextSym.Case <> esriTextCase.esriTCNormal Then
            mxdProps.bTextCase = True
            If pFormTextSym.Case = esriTextCase.esriTCAllCaps Then sw.WriteLine(InsertTabs(lTabLevel) & "All caps")
            If pFormTextSym.Case = esriTextCase.esriTCLowercase Then sw.WriteLine(InsertTabs(lTabLevel) & "Lower case")
            If pFormTextSym.Case = esriTextCase.esriTCSmallCaps Then sw.WriteLine(InsertTabs(lTabLevel) & "Small caps")
        End If
        If Math.Abs(pFormTextSym.CharacterSpacing) > 0 Then
            mxdProps.bCharSpacing = True
            sw.WriteLine(InsertTabs(lTabLevel) & "Char spacing: " & pFormTextSym.CharacterSpacing)
        End If
        If Math.Abs(pFormTextSym.Leading) > 0 Then
            mxdProps.bLeading = True
            sw.WriteLine(InsertTabs(lTabLevel) & "Leading: " & pFormTextSym.Leading)
        End If
        If Math.Abs(pFormTextSym.CharacterWidth) > 100 Then
            mxdProps.bCharWidth = True
            sw.WriteLine(InsertTabs(lTabLevel) & "Char width: " & pFormTextSym.CharacterWidth)
        End If
        If Math.Abs(pFormTextSym.WordSpacing) > 100 Then
            mxdProps.bWordSpacing = True
            sw.WriteLine(InsertTabs(lTabLevel) & "Word spacing: " & pFormTextSym.WordSpacing)
        End If
        If Not pFormTextSym.Kerning Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Kerning off")
            mxdProps.bKerningOff = True
        End If
        'Advanced Text tab
        If Not pFormTextSym.FillSymbol Is Nothing Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Text fill pattern:")
            mxdProps.bFillSymbol = True
            GetSymbolProps(pFormTextSym.FillSymbol, lTabLevel + 1, False)
        End If
        If Not pTextBackground Is Nothing Then
            mxdProps.bTextBackground = True
            sw.WriteLine(InsertTabs(lTabLevel) & "Text background:")
            If TypeOf pTextBackground Is IBalloonCallout Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Balloon callout")
                pBalloonCallout = pTextBackground
                sw.WriteLine(InsertTabs(lTabLevel + 2) & "Tolerance: " & pBalloonCallout.LeaderTolerance())
                mxdProps.bBalloonCallout = True
            ElseIf TypeOf pTextBackground Is ILineCallout Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Line callout")
                pLineCallout = pTextBackground
                sw.WriteLine(InsertTabs(lTabLevel + 2) & "Gap: " & pLineCallout.Gap)
                sw.WriteLine(InsertTabs(lTabLevel + 2) & "Tolerance: " & pLineCallout.LeaderTolerance())
                If Not pLineCallout.LeaderLine Is Nothing Then sw.WriteLine(InsertTabs(lTabLevel + 2) & "Leader line")
                If Not pLineCallout.AccentBar Is Nothing Then sw.WriteLine(InsertTabs(lTabLevel + 2) & "Accent bar")
                If Not pLineCallout.Border Is Nothing Then sw.WriteLine(InsertTabs(lTabLevel + 2) & "Border")
                sw.WriteLine(InsertTabs(lTabLevel + 2) & "Callout style: " & GetLineCalloutStyle(pLineCallout.Style))
                mxdProps.bLineCallout = True
            ElseIf TypeOf pTextBackground Is ISimpleLineCallout Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Simple line callout")
                pSimpleLineCallout = pTextBackground
                sw.WriteLine(InsertTabs(lTabLevel + 2) & "Tolerance: " & pSimpleLineCallout.LeaderTolerance())
                If pSimpleLineCallout.AutoSnap Then sw.WriteLine(InsertTabs(lTabLevel + 2) & "Snap to text")
                mxdProps.bSimpleLineCallout = True
            ElseIf TypeOf pTextBackground Is IMarkerTextBackground Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Marker text background")
                pMarkerTextBkg = pTextBackground
                pMarkerSymbol = pMarkerTextBkg.Symbol
                mxdProps.bMarkerTextBkg = True
                If pMarkerTextBkg.ScaleToFit Then
                    sw.WriteLine(InsertTabs(lTabLevel + 2) & "Scale to fit")
                    mxdProps.bScaletoFit = True
                Else
                    sw.WriteLine(InsertTabs(lTabLevel + 2) & "Size: " & pMarkerSymbol.Size)
                End If
            End If
        End If 'Text background
        If Math.Abs(pFormTextSym.ShadowXOffset) > 0 And Math.Abs(pFormTextSym.ShadowYOffset) > 0 Then mxdProps.bShadow = True
        If Math.Abs(pFormTextSym.ShadowXOffset) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Shadow X offset: " & pFormTextSym.ShadowXOffset)
        If Math.Abs(pFormTextSym.ShadowYOffset) > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Shadow Y offset: " & pFormTextSym.ShadowYOffset)
        If Math.Abs(pFormTextSym.ShadowYOffset) > 0 Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Color (RGB): " & GetRGB((pFormTextSym.ShadowColor)))
            sw.WriteLine(InsertTabs(lTabLevel) & "Color (CMYK): " & GetCMYK((pFormTextSym.ShadowColor)))
        End If
        'Mask tab
        If pMask.MaskStyle = esriMaskStyle.esriMSHalo Then
            mxdProps.bHalo = True
            pMaskSym = pMask.MaskSymbol
            sw.WriteLine(InsertTabs(lTabLevel) & "Halo")
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Size: " & pMask.MaskSize)
            GetSymbolProps(pMaskSym, lTabLevel + 1, False)
            'If Not pMaskSym Is Nothing Then
            '  If pMaskSym.Color Is Nothing Then
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "No fill color")
            '  Else
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Color (RGB): " & GetRGB((pMaskSym.Color)))
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Color (CMYK): " & GetCMYK((pMaskSym.Color)))
            '  End If
            '  sw.WriteLine(InsertTabs(lTabLevel + 1) & "Outline width: " & pMaskSym.Outline.Width)
            '  sw.WriteLine(InsertTabs(lTabLevel + 1) & "Outline color (RGB): " & GetRGB(pMaskSym.Outline.Color))
            '  sw.WriteLine(InsertTabs(lTabLevel + 1) & "Outline color (CMYK): " & GetCMYK(pMaskSym.Outline.Color))
            '  If TypeOf pMaskSym Is ISimpleFillSymbol Then
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Simple fill symbol")
            '  ElseIf TypeOf pMaskSym Is IPictureFillSymbol Then
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Picture fill symbol")
            '  ElseIf TypeOf pMaskSym Is IMarkerFillSymbol Then
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Marker fill symbol")
            '  ElseIf TypeOf pMaskSym Is ILineFillSymbol Then
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Line fill symbol")
            '  ElseIf TypeOf pMaskSym Is IGradientFillSymbol Then
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Gradient fill symbol")
            '  ElseIf TypeOf pMaskSym Is ITextureFillSymbol Then
            '    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Texture fill symbol")
            '  End If
            'End If
        End If
        'Misc
        If pCharOrientation.CJKCharactersRotation Then
            mxdProps.bCJK = True
            sw.WriteLine(InsertTabs(lTabLevel) & "CJK character rotation")
        End If
        sw.WriteLine(InsertTabs(lTabLevel) & "Flip angle: " & pFormTextSym.FlipAngle)
        If pFormTextSym.TypeSetting Then sw.WriteLine(InsertTabs(lTabLevel) & "Typesetting")
        If pFormTextSym.Direction = esriTextDirection.esriTDAngle Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Text angled at: " & pFormTextSym.Angle())
        ElseIf pFormTextSym.Direction = esriTextDirection.esriTDHorizontal Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal text")
        ElseIf pFormTextSym.Direction = esriTextDirection.esriTDVertical Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Vertical text")
        End If
        If pSimpleTextSymbol.BreakCharacter > 0 Then sw.WriteLine(InsertTabs(lTabLevel) & "Break character: " & Chr(pSimpleTextSymbol.BreakCharacter))
        If pSimpleTextSymbol.Clip Then sw.WriteLine(InsertTabs(lTabLevel) & "Clipped per geometry")

    End Sub

    Sub GetMLEProps(ByRef pLabelEngineLayerProperties As ILabelEngineLayerProperties2, _
                    ByRef pTextSym As ITextSymbol, ByRef lTabLevel As Integer)
        sw.Flush()
        Dim i As Integer
        Dim sTmp As String
        Dim pMpxOpLProps As IMaplexOverposterLayerProperties
        Dim pMpxStackProps As IMaplexLabelStackingProperties
        Dim pMpxOffAlongLine As IMaplexOffsetAlongLineProperties
        Dim pMpxRotProps As IMaplexRotationProperties
        Dim pFormTextSym As IFormattedTextSymbol
        Dim pMpxOpLProps2 As IMaplexOverposterLayerProperties2 = Nothing
        Dim pMpxRotProps2 As IMaplexRotationProperties2 = Nothing
        Dim pMpxOpLProps3 As IMaplexOverposterLayerProperties3 = Nothing
        Dim pMpxOpLProps4 As IMaplexOverposterLayerProperties4 = Nothing
        Dim bIsStreetLine, bIsOffsetLine, bIsRegularLine As Boolean
        Dim bIsStreetAddress, bIsRegularPolygon As Boolean

        'layer file - don't know if MLE or SLE
        Try
            pMpxOpLProps = pLabelEngineLayerProperties.OverposterLayerProperties
            pMpxStackProps = pMpxOpLProps.LabelStackingProperties
            pMpxOffAlongLine = pMpxOpLProps.OffsetAlongLineProperties
            pMpxRotProps = pMpxOpLProps.RotationProperties
            pFormTextSym = pTextSym
            If m_Version >= 93 Then
                pMpxOpLProps2 = pLabelEngineLayerProperties.OverposterLayerProperties
                pMpxRotProps2 = pMpxOpLProps.RotationProperties
            End If '>=9.3
            If m_Version >= 94 Then
                pMpxOpLProps3 = pLabelEngineLayerProperties.OverposterLayerProperties
            End If '>=9.4
            If m_Version >= 101 Then
                pMpxOpLProps4 = pLabelEngineLayerProperties.OverposterLayerProperties
            End If '>=10.1
        Catch ex As Exception
            Return
        End Try

        sw.WriteLine(InsertTabs(lTabLevel - 1) & "Label properties:")
        'Placement tab
        If pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPoint Then
            'points
            If pMpxOpLProps.PointPlacementMethod Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Fixed position: " & GetPointPlacementMethod(pMpxOpLProps.PointPlacementMethod))
                mxdProps.bPointFixed = True
                If pMpxOpLProps.CanShiftPointLabel Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "May shift")
                    mxdProps.bMayShift = True
                End If
            Else
                sw.WriteLine(InsertTabs(lTabLevel) & "Best position")
                mxdProps.bPointBest = True
                If pMpxOpLProps.EnablePointPlacementPriorities Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Point zones")
                    mxdProps.bPointZones = True
                    sTmp = pMpxOpLProps.PointPlacementPriorities.AboveCenter & _
                        " " & pMpxOpLProps.PointPlacementPriorities.AboveRight & " " & pMpxOpLProps.PointPlacementPriorities.CenterRight & _
                        " " & pMpxOpLProps.PointPlacementPriorities.BelowRight & " " & pMpxOpLProps.PointPlacementPriorities.BelowCenter & _
                        " " & pMpxOpLProps.PointPlacementPriorities.BelowLeft & " " & pMpxOpLProps.PointPlacementPriorities.CenterLeft & _
                        " " & pMpxOpLProps.PointPlacementPriorities.AboveLeft
                    If Not sTmp.Equals("2 1 3 5 7 8 6 4") Then mxdProps.bAlteredZones = True
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Priorities (clockwise from north): " & sTmp)
                End If
            End If
            If Not pMpxOpLProps.PointPlacementMethod = esriMaplexPointPlacementMethod.esriMaplexCenteredOnPoint Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Offset distance: " & pMpxOpLProps.PrimaryOffset & " " & GetMaplexUnits((pMpxOpLProps.PrimaryOffsetUnit)))
                sw.WriteLine(InsertTabs(lTabLevel) & "Maximum offset: " & pMpxOpLProps.SecondaryOffset & "%")
                If pMpxOpLProps.PrimaryOffset Then mxdProps.bPointOffDist = True
                If Math.Abs(pMpxOpLProps.SecondaryOffset) - 100 > 0 Then mxdProps.bPointMaxOffset = True
                If m_Version >= 93 Then
                    If pMpxOpLProps2.IsOffsetFromFeatureGeometry Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Use feature geometry")
                        mxdProps.bPointFtrGeom = True
                    End If
                End If '>=9.3
                If m_Version >= 101 Then
                    If pMpxOpLProps4.UseExactSymbolOutline Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Use symbol outline")
                        mxdProps.bSymbolOutline = True
                    End If
                End If '>=10.1
            End If
            If pMpxOpLProps.GraticuleAlignment Then
                If m_Version >= 93 Then
                    Select Case pMpxOpLProps2.GraticuleAlignmentType
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGACurved
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment curved")
                            mxdProps.bPointGACurv = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGACurvedNoFlip
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment curved (no flip)")
                            mxdProps.bPointGACurvNoFlip = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGAStraight
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment straight")
                            mxdProps.bPointGAStr = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGAStraightNoFlip
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment straight (no flip)")
                            mxdProps.bPointGAStrNoFlip = True
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: graticule alignment not known")
                    End Select
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment")
                    mxdProps.bPointGAStr = True
                End If '>=9.2
            End If
            If pMpxRotProps.Enable Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Point rotation")
                mxdProps.bPointRotation = True
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Field: " & pMpxRotProps.RotationField)
                If m_Version >= 93 Then
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Additional angle: " & pMpxRotProps2.AdditionalAngle)
                    If pMpxRotProps2.AdditionalAngle <> 0 Then mxdProps.bPointRotAngle = True
                End If '>=9.3
                If pMpxRotProps.RotationType = esriLabelRotationType.esriRotateLabelArithmetic Then
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Arithmetic")
                Else
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Geographic")
                End If
                If m_Version >= 93 Then
                    Select Case pMpxRotProps2.AlignmentType
                        Case esriMaplexRotationAlignmentType.esriMaplexRotationAlignmentHorizontal
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Horizontal alignment")
                        Case esriMaplexRotationAlignmentType.esriMaplexRotationAlignmentPerpendicular
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Perpendicular alignment")
                        Case esriMaplexRotationAlignmentType.esriMaplexRotationAlignmentStraight
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Straight alignment")
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Error: rotation alignment not known")
                    End Select
                End If '>=9.3
                If Not pMpxRotProps.AlignLabelToAngle Then
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "May flip")
                    mxdProps.bPointRotFlip = True
                End If
            End If

        ElseIf pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolyline Then
            'lines
            Select Case pMpxOpLProps.LinePlacementMethod
                Case esriMaplexLinePlacementMethod.esriMaplexCenteredCurvedOnLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line centered curved")
                    mxdProps.bLineCenCur = True
                Case esriMaplexLinePlacementMethod.esriMaplexCenteredHorizontalOnLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line centered horizontal")
                    mxdProps.bLineCenHor = True
                Case esriMaplexLinePlacementMethod.esriMaplexCenteredPerpendicularOnLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line centered perpendicular")
                    mxdProps.bLineCenPer = True
                Case esriMaplexLinePlacementMethod.esriMaplexCenteredStraightOnLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line centered straight")
                    mxdProps.bLineCenStr = True
                Case esriMaplexLinePlacementMethod.esriMaplexOffsetCurvedFromLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line offset curved")
                    mxdProps.bLineOffCur = True
                    bIsOffsetLine = True
                Case esriMaplexLinePlacementMethod.esriMaplexOffsetHorizontalFromLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line offset horizontal")
                    mxdProps.bLineOffHor = True
                    bIsOffsetLine = True
                Case esriMaplexLinePlacementMethod.esriMaplexOffsetPerpendicularFromLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line offset perpendicular")
                    mxdProps.bLineOffPer = True
                    bIsOffsetLine = True
                Case esriMaplexLinePlacementMethod.esriMaplexOffsetStraightFromLine
                    sw.WriteLine(InsertTabs(lTabLevel) & "Line offset straight")
                    mxdProps.bLineOffStr = True
                    bIsOffsetLine = True
                Case Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Error: line placement method not known")
            End Select
            If m_Version >= 93 Then
                Select Case pMpxOpLProps2.LineFeatureType
                    Case esriMaplexLineFeatureType.esriMaplexContourFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "Contour")
                        mxdProps.bLineContour = True
                    Case esriMaplexLineFeatureType.esriMaplexLineFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "Regular")
                        mxdProps.bLineRegular = True
                        bIsRegularLine = True
                    Case esriMaplexLineFeatureType.esriMaplexRiverFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "River")
                        mxdProps.bLineRiver = True
                    Case esriMaplexLineFeatureType.esriMaplexStreetAddressRange
                        sw.WriteLine(InsertTabs(lTabLevel) & "Street address")
                        mxdProps.bLineStreetAdd = True
                        bIsStreetAddress = True
                    Case esriMaplexLineFeatureType.esriMaplexStreetFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "Street")
                        mxdProps.bLineStreet = True
                        bIsStreetLine = True
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Error: line feature type not known")
                End Select
            Else
                If pMpxOpLProps.IsStreetPlacement Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Street")
                    mxdProps.bLineStreet = True
                    bIsStreetLine = True
                Else
                    mxdProps.bLineRegular = True
                    bIsRegularLine = True
                End If
            End If '>=9.3
            'offset from line
            If bIsOffsetLine Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Offset distance: " & pMpxOpLProps.PrimaryOffset & " " & _
                             GetMaplexUnits((pMpxOpLProps.PrimaryOffsetUnit)))
                'note: no preferred offset for lines
                If pMpxOpLProps.PrimaryOffset Then mxdProps.bLineOffDist = True
                If Not bIsStreetLine Then
                    Select Case pMpxOpLProps.ConstrainOffset
                        Case esriMaplexConstrainOffset.esriMaplexAboveLine
                            sw.WriteLine(InsertTabs(lTabLevel) & "Constrain above")
                            mxdProps.bConstrainAbove = True
                        Case esriMaplexConstrainOffset.esriMaplexBelowLine
                            sw.WriteLine(InsertTabs(lTabLevel) & "Constrain below")
                            mxdProps.bConstrainBelow = True
                        Case esriMaplexConstrainOffset.esriMaplexLeftOfLine
                            sw.WriteLine(InsertTabs(lTabLevel) & "Constrain to left")
                            mxdProps.bConstrainLeft = True
                        Case esriMaplexConstrainOffset.esriMaplexNoConstraint
                            sw.WriteLine(InsertTabs(lTabLevel) & "No constraint")
                            mxdProps.bNoConstraint = True
                        Case esriMaplexConstrainOffset.esriMaplexRightOfLine
                            sw.WriteLine(InsertTabs(lTabLevel) & "Constrain to right")
                            mxdProps.bConstrainRight = True
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: offset constraint not known")
                    End Select
                End If
                If m_Version >= 93 Then
                    If pMpxOpLProps2.IsOffsetFromFeatureGeometry Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Use feature geometry")
                        mxdProps.bLineFtrGeom = True
                    End If
                End If '>=9.3
            End If
            'offset along line
            If bIsRegularLine Or bIsStreetAddress Then
                If pMpxOffAlongLine.PlacementMethod = esriMaplexOffsetAlongLineMethod.esriMaplexBestPositionAlongLine Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Best position along line")
                    mxdProps.bLineBestAlong = True
                Else
                    Select Case pMpxOffAlongLine.PlacementMethod
                        Case esriMaplexOffsetAlongLineMethod.esriMaplexAfterEndOfLine
                            sw.WriteLine(InsertTabs(lTabLevel) & "After end of line")
                            mxdProps.bLineAfterEnd = True
                        Case esriMaplexOffsetAlongLineMethod.esriMaplexAlongLineFromEnd
                            sw.WriteLine(InsertTabs(lTabLevel) & "Along from end of line")
                            mxdProps.bLineFromEnd = True
                        Case esriMaplexOffsetAlongLineMethod.esriMaplexAlongLineFromStart
                            sw.WriteLine(InsertTabs(lTabLevel) & "Along from start of line")
                            mxdProps.bLineFromStart = True
                        Case esriMaplexOffsetAlongLineMethod.esriMaplexBeforeStartOfLine
                            sw.WriteLine(InsertTabs(lTabLevel) & "Before start of line")
                            mxdProps.bLineBeforeStart = True
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: line placement method not known")
                    End Select
                    Select Case pMpxOffAlongLine.LabelAnchorPoint
                        Case esriMaplexLabelAnchorPoint.esriMaplexCenterOfLabel
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Center of label")
                        Case esriMaplexLabelAnchorPoint.esriMaplexFurthestSideOfLabel
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Furthest side of label")
                        Case esriMaplexLabelAnchorPoint.esriMaplexNearestSideOfLabel
                            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Nearest side of label")
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: label anchor point not known")
                    End Select
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Distance: " & pMpxOffAlongLine.Distance & " " & _
                                 GetMaplexUnits((pMpxOffAlongLine.DistanceUnit)))
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Tolerance: " & pMpxOffAlongLine.Tolerance)
                    If pMpxOffAlongLine.UseLineDirection Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Use line direction")
                End If
            End If 'offset
            If m_Version >= 101 Then
                If pMpxOpLProps4.AllowStraddleStacking And pMpxOpLProps2.LineFeatureType = esriMaplexLineFeatureType.esriMaplexLineFeature Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Allow straddle stacking")
                    mxdProps.bStraddlacking = True
                End If
            End If
            If m_Version >= 93 Then
                If pMpxOpLProps2.EnableSecondaryOffset Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Secondary offset between " & pMpxOpLProps2.SecondaryOffsetMinimum & _
                                 " and " & pMpxOpLProps2.SecondaryOffsetMaximum & " " & GetMaplexUnits((pMpxOpLProps.PrimaryOffsetUnit)))
                    mxdProps.bLineSecOff = True
                End If
                If pMpxOpLProps2.LineFeatureType = esriMaplexLineFeatureType.esriMaplexContourFeature Then
                    If pMpxOpLProps2.ContourAlignmentType = esriMaplexContourAlignmentType.esriMaplexPageAlignment Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Contour page aligned at " & pMpxOpLProps2.ContourMaximumAngle)
                        mxdProps.bContourPage = True
                    Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Contour uphill")
                        mxdProps.bContourUphill = True
                    End If
                    If pMpxOpLProps2.ContourLadderType = esriMaplexContourLadderType.esriMaplexNoLadder Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "No ladders")
                        mxdProps.bContourNoLadder = True
                    Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Ladders")
                        mxdProps.bContourLadder = True
                    End If
                End If
            End If '>=9.3
            If bIsStreetLine Then
                If m_Version >= 93 Then
                    If pMpxOpLProps2.CanPlaceLabelOnTopOfFeature Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Street horizontal")
                        mxdProps.bStreetHorz = True
                    End If
                    If pMpxOpLProps2.CanReduceLeading Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Street reduce leading")
                        mxdProps.bStreetReduce = True
                    End If
                    If pMpxOpLProps2.CanFlipStackedStreetLabel Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Street primary name under")
                        mxdProps.bStreetPrimary = True
                    End If
                End If '>=9.3
                If pMpxOpLProps.SpreadWords Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Street spread words")
                    mxdProps.bStreetSpread = True
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Maximum spacing: " & pMpxOpLProps.MaximumWordSpacing & "%")
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Default spacing: " & pFormTextSym.WordSpacing & "%")
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Pref clearance: " & pMpxOpLProps.PreferredEndOfStreetClearance & "%")
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Min clearance: " & pMpxOpLProps.MinimumEndOfStreetClearance & "%")
                End If
            End If
            If pMpxOpLProps.AlignLabelToLineDirection Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Align to line direction")
                mxdProps.bLineDirection = True
            End If
            If pMpxOpLProps.RepeatLabel Then
                sTmp = InsertTabs(lTabLevel) & "Repeat after: " & pMpxOpLProps.MinimumRepetitionInterval
                If m_Version >= 93 Then
                    sTmp = sTmp & " " & GetMaplexUnits((pMpxOpLProps2.RepetitionIntervalUnit))
                Else
                    sTmp = sTmp & " Map Units"
                End If
                sw.WriteLine(sTmp)
                mxdProps.bLineRepeat = True
                If m_Version >= 101 Then
                    If pMpxOpLProps4.PreferLabelNearJunction Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Prefer label near junction")
                        sw.WriteLine(InsertTabs(lTabLevel) & "Clearance: " & pMpxOpLProps4.PreferLabelNearJunctionClearance)
                        mxdProps.bLabelNearJunction = True
                    End If
                    If pMpxOpLProps4.PreferLabelNearMapBorder Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Prefer label near border")
                        sw.WriteLine(InsertTabs(lTabLevel) & "Clearance: " & pMpxOpLProps4.PreferLabelNearMapBorderClearance)
                        mxdProps.bLabelNearBorder = True
                    End If
                End If '10.1
            End If
            If pMpxOpLProps.SpreadCharacters Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Spread characters: " & pMpxOpLProps.MaximumCharacterSpacing & "%" & _
                             " (default " & pFormTextSym.CharacterSpacing & "%)")
                mxdProps.bLineSpread = True
            End If
            'graticule alignment - horizontal labels only
            If pMpxOpLProps.GraticuleAlignment And ( _
                pMpxOpLProps.LinePlacementMethod = esriMaplexLinePlacementMethod.esriMaplexCenteredHorizontalOnLine Or _
                pMpxOpLProps.LinePlacementMethod = esriMaplexLinePlacementMethod.esriMaplexOffsetHorizontalFromLine Or _
                pMpxOpLProps.PreferHorizontalPlacement) Then
                If m_Version >= 93 Then
                    Select Case pMpxOpLProps2.GraticuleAlignmentType
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGACurved
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment curved")
                            mxdProps.bLineGACurv = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGACurvedNoFlip
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment curved (no flip)")
                            mxdProps.bLineGACurvNoFlip = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGAStraight
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment straight")
                            mxdProps.bLineGAStr = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGAStraightNoFlip
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment straight (no flip)")
                            mxdProps.bLineGAStrNoFlip = True
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: graticule alignment not known")
                    End Select
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment")
                    mxdProps.bLineGAStr = True
                End If '>=9.3
            End If 'graticule
            If m_Version >= 101 Then
                If pMpxOpLProps4.EnableConnection Then
                    Select Case pMpxOpLProps4.ConnectionType
                        Case esriMaplexConnectionType.esriMaplexMinimizeLabels
                            sw.WriteLine(InsertTabs(lTabLevel) & "Line connection: Minimize")
                            mxdProps.bMinimize = True
                        Case esriMaplexConnectionType.esriMaplexUnambiguous
                            sw.WriteLine(InsertTabs(lTabLevel) & "Line connection: Unambiguous")
                            mxdProps.bUnambiguous = True
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Line connection: Unknown")
                    End Select
                Else
                    Select Case pMpxOpLProps4.MultiPartOption
                        Case esriMaplexMultiPartOption.esriMaplexOneLabelPerFeature
                            sw.WriteLine(InsertTabs(lTabLevel) & "One label per feature")
                            mxdProps.bMultiOptionFeature = True
                        Case esriMaplexMultiPartOption.esriMaplexOneLabelPerPart
                            sw.WriteLine(InsertTabs(lTabLevel) & "One label per feature part")
                            mxdProps.bMultiOptionPart = True
                        Case esriMaplexMultiPartOption.esriMaplexOneLabelPerSegment
                            sw.WriteLine(InsertTabs(lTabLevel) & "One label per segment")
                            mxdProps.bMultiOptionSegment = True
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: Unknown multi-part option")
                    End Select
                End If 'connection
            End If '10.1

        ElseIf pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon Then
            'polygons
            If m_Version >= 93 Then
                Select Case pMpxOpLProps2.PolygonFeatureType
                    Case esriMaplexPolygonFeatureType.esriMaplexPolygonFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "Regular")
                        mxdProps.bPolyRegular = True
                        bIsRegularPolygon = True
                    Case esriMaplexPolygonFeatureType.esriMaplexLandParcelFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "Land parcel")
                        mxdProps.bPolyParcel = True
                    Case esriMaplexPolygonFeatureType.esriMaplexRiverPolygonFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "River")
                        mxdProps.bPolyRiver = True
                    Case esriMaplexPolygonFeatureType.esriMaplexPolygonBoundaryFeature
                        sw.WriteLine(InsertTabs(lTabLevel) & "Boundary feature type")
                        mxdProps.bPolyBdy = True
                        If m_Version >= 94 Then
                            If pMpxOpLProps3.BoundaryLabelingAllowSingleSided Then
                                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Single sided")
                                mxdProps.bPolyBdySingleSided = True
                            End If
                            If pMpxOpLProps3.BoundaryLabelingAllowHoles Then
                                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Allow holes")
                                mxdProps.bPolyBdyAllowHoles = True
                            End If
                            If pMpxOpLProps3.BoundaryLabelingSingleSidedOnLine Then
                                sw.WriteLine(InsertTabs(lTabLevel + 1) & "On line")
                                mxdProps.bPolyBdyOnLine = True
                            End If
                        End If '9.4 (10.0)
                    Case Else
                        sw.WriteLine(InsertTabs(lTabLevel) & "Error: polygon feature type not known")
                End Select
            Else
                If pMpxOpLProps.LandParcelPlacement Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Land parcel")
                    mxdProps.bPolyParcel = True
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Regular")
                    mxdProps.bPolyRegular = True
                    bIsRegularPolygon = True
                End If
            End If '>=9.3
            Select Case pMpxOpLProps.PolygonPlacementMethod
                Case esriMaplexPolygonPlacementMethod.esriMaplexCurvedAroundPolygon
                    sw.WriteLine(InsertTabs(lTabLevel) & "Offset curved")
                    mxdProps.bPolyOffCurv = True
                Case esriMaplexPolygonPlacementMethod.esriMaplexCurvedInPolygon
                    sw.WriteLine(InsertTabs(lTabLevel) & "Curved")
                    mxdProps.bPolyCurv = True
                Case esriMaplexPolygonPlacementMethod.esriMaplexHorizontalAroundPolygon
                    sw.WriteLine(InsertTabs(lTabLevel) & "Offset horizontal")
                    mxdProps.bPolyOffHorz = True
                Case esriMaplexPolygonPlacementMethod.esriMaplexHorizontalInPolygon
                    sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal")
                    mxdProps.bPolyHorz = True
                Case esriMaplexPolygonPlacementMethod.esriMaplexRepeatAlongBoundary
                    sw.WriteLine(InsertTabs(lTabLevel) & "Boundary")
                    mxdProps.bPolyBdy = True
                Case esriMaplexPolygonPlacementMethod.esriMaplexStraightInPolygon
                    sw.WriteLine(InsertTabs(lTabLevel) & "Straight")
                    mxdProps.bPolyStr = True
                Case Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Error: polygon placement method not known")
            End Select
            If bIsRegularPolygon And _
              (pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexStraightInPolygon Or _
               pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexCurvedInPolygon) And _
               pMpxOpLProps.PreferHorizontalPlacement Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Try horizontal first")
                mxdProps.bPolyTryHorz = True
            End If
            If pMpxOpLProps.CanPlaceLabelOutsidePolygon Then
                sw.WriteLine(InsertTabs(lTabLevel) & "May place outside")
                mxdProps.bPolyMayPlaceOutside = True
            End If
            If pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexCurvedAroundPolygon Or _
                pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexHorizontalAroundPolygon Or _
                pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexRepeatAlongBoundary Or _
                pMpxOpLProps.CanPlaceLabelOutsidePolygon Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Offset distance: " & pMpxOpLProps.PrimaryOffset & " " & GetMaplexUnits((pMpxOpLProps.PrimaryOffsetUnit)))
                sw.WriteLine(InsertTabs(lTabLevel) & "Maximum offset: " & pMpxOpLProps.SecondaryOffset & "%")
                If pMpxOpLProps.PrimaryOffset Then mxdProps.bPolyOffDist = True
                If Math.Abs(pMpxOpLProps.SecondaryOffset) - 100 > 0 Then mxdProps.bPolyMaxOffset = True
                If m_Version >= 93 Then
                    If pMpxOpLProps2.IsOffsetFromFeatureGeometry Then
                        sw.WriteLine(InsertTabs(lTabLevel) & "Use feature geometry")
                        mxdProps.bPolyFtrGeom = True
                    End If
                End If '>=9.3
            End If
            'graticule alignment - horizontal labels only
            If pMpxOpLProps.GraticuleAlignment And ( _
                pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexHorizontalAroundPolygon Or _
                pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexHorizontalInPolygon Or _
                pMpxOpLProps.PreferHorizontalPlacement) Then
                If m_Version >= 93 Then
                    Select Case pMpxOpLProps2.GraticuleAlignmentType
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGACurved
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment curved")
                            mxdProps.bPolyGACurv = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGACurvedNoFlip
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment curved (no flip)")
                            mxdProps.bPolyGACurvNoFlip = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGAStraight
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment straight")
                            mxdProps.bPolyGAStr = True
                        Case esriMaplexGraticuleAlignmentType.esriMaplexGAStraightNoFlip
                            sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment straight (no flip)")
                            mxdProps.bPolyGAStrNoFlip = True
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: graticule alignment not known")
                    End Select
                Else
                    sw.WriteLine(InsertTabs(lTabLevel) & "Graticule alignment")
                    mxdProps.bPolyGAStr = True
                End If '>=9.2
            End If
            If m_Version >= 93 Then
                If pMpxOpLProps2.EnablePolygonFixedPosition Then
                    sTmp = ""
                    For i = 0 To 8
                        sTmp = sTmp & pMpxOpLProps2.PolygonInternalZones(i) & " "
                    Next
                    sw.WriteLine(InsertTabs(lTabLevel) & "Internal zones: " & sTmp)
                    mxdProps.bPolyIntZones = True
                End If
                If pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexCurvedAroundPolygon Or _
                    pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexHorizontalAroundPolygon Or _
                    pMpxOpLProps.CanPlaceLabelOutsidePolygon Then
                    'offset label: ext zones and anchor points
                    sTmp = ""
                    For i = 0 To 7
                        sTmp = sTmp & pMpxOpLProps2.PolygonExternalZones(i) & " "
                    Next
                    sw.WriteLine(InsertTabs(lTabLevel) & "External zones: " & sTmp)
                    mxdProps.bPolyExtZones = True
                    Select Case pMpxOpLProps2.PolygonAnchorPointType
                        Case esriMaplexAnchorPointType.esriMaplexErodedCenter
                            sw.WriteLine(InsertTabs(lTabLevel) & "Eroded center")
                        Case esriMaplexAnchorPointType.esriMaplexGeometricCenter
                            sw.WriteLine(InsertTabs(lTabLevel) & "Geometric center")
                        Case esriMaplexAnchorPointType.esriMaplexPerimeter
                            sw.WriteLine(InsertTabs(lTabLevel) & "Closest point on boundary")
                        Case esriMaplexAnchorPointType.esriMaplexUnclippedGeometricCenter
                            sw.WriteLine(InsertTabs(lTabLevel) & "Unclipped geometric center")
                        Case Else
                            sw.WriteLine(InsertTabs(lTabLevel) & "Error: anchor point type not known")
                    End Select
                    If pMpxOpLProps2.PolygonAnchorPointType <> esriMaplexAnchorPointType.esriMaplexGeometricCenter Then
                        mxdProps.bPolyAnchor = True
                    End If
                End If
            End If '>=9.3
            If pMpxOpLProps.RepeatLabel Then
                If pMpxOpLProps.PolygonPlacementMethod = esriMaplexPolygonPlacementMethod.esriMaplexRepeatAlongBoundary Or _
                    pMpxOpLProps2.PolygonFeatureType = esriMaplexPolygonFeatureType.esriMaplexPolygonBoundaryFeature Then _
                    sTmp = InsertTabs(lTabLevel) & "Repeat along boundary: " Else _
                    sTmp = InsertTabs(lTabLevel) & "Repeat label: "
                sTmp = sTmp & pMpxOpLProps.MinimumRepetitionInterval
                If m_Version >= 93 Then
                    sTmp = sTmp & " " & GetMaplexUnits((pMpxOpLProps2.RepetitionIntervalUnit))
                Else
                    sTmp = sTmp & " Map Units"
                End If
                sw.WriteLine(sTmp)
                mxdProps.bPolyRepeat = True
            End If 'repeat
            If pMpxOpLProps.SpreadCharacters Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Spread characters: " & pMpxOpLProps.MaximumCharacterSpacing & "%" & _
                             " (default " & pFormTextSym.CharacterSpacing & "%)")
                mxdProps.bPolySpread = True
            End If
            'holes
            If m_Version >= 94 Then
                If Not pMpxOpLProps3.AvoidPolygonHoles Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Allow holes")
                    mxdProps.bPolyAllowHoles = True
                End If
            End If
            If m_Version >= 101 Then
                If pMpxOpLProps4.LabelLargestPolygon Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Label largest polygon")
                    mxdProps.bLargestOnly = True
                End If
            End If '10.1
        End If 'feature type
        If m_Version >= 101 Then
            If Not pMpxOpLProps4.RemoveExtraWhiteSpace Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Remove extra whitespace is off")
                mxdProps.bWhitespace = True
            End If
            If pMpxOpLProps4.RemoveExtraLineBreaks Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Remove extra line breaks")
                mxdProps.bLinebreaks = True
            End If
        End If

        'Strategy tab
        sw.WriteLine(InsertTabs(lTabLevel - 1) & "Strategies:")
        If pMpxOpLProps.CanStackLabel Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Stacking")
            mxdProps.bStack = True
            If pMpxStackProps.StackJustification = esriMaplexStackingJustification.esriMaplexChooseBestJustification Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Best justification")
            ElseIf pMpxStackProps.StackJustification = esriMaplexStackingJustification.esriMaplexConstrainJustificationCenter Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Center justification")
                mxdProps.bStackC = True
            ElseIf pMpxStackProps.StackJustification = esriMaplexStackingJustification.esriMaplexConstrainJustificationLeftOrRight Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Left or right justification")
                mxdProps.bStackLorR = True
            ElseIf pMpxStackProps.StackJustification = esriMaplexStackingJustification.esriMaplexConstrainJustificationLeft Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Left justification")
                mxdProps.bStackL = True
            ElseIf pMpxStackProps.StackJustification = esriMaplexStackingJustification.esriMaplexConstrainJustificationRight Then
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Right justification")
                mxdProps.bStackR = True
            End If
            If pMpxStackProps.SeparatorCount Then 'mxdProps.lSeparators < 32 And 
                For i = 0 To pMpxStackProps.SeparatorCount - 1
                    pMpxStackProps.QuerySeparator(i, mxdProps.sSeparators(mxdProps.lSeparators), mxdProps.bSepVis, mxdProps.bSepFor, mxdProps.bSepAft)
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Separator '" & mxdProps.sSeparators(mxdProps.lSeparators) & "' Visible: " & mxdProps.bSepVis & _
                                 " Force split: " & mxdProps.bSepFor & " Split after: " & mxdProps.bSepAft)
                    AddIfUnique(mxdProps.lSeparators, mxdProps.sSeparators, ARRAY_SIZE)
                Next
            End If
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Max lines: " & pMpxStackProps.MaximumNumberOfLines)
            If pMpxStackProps.MaximumNumberOfLines <> 3 Then mxdProps.bMaxLines = True
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Min chars per line: " & pMpxStackProps.MinimumNumberOfCharsPerLine)
            If pMpxStackProps.MinimumNumberOfCharsPerLine <> 3 Then mxdProps.bMinChars = True
            sw.WriteLine(InsertTabs(lTabLevel + 1) & "Max chars per line: " & pMpxStackProps.MaximumNumberOfCharsPerLine)
            If pMpxStackProps.MaximumNumberOfCharsPerLine <> 24 Then mxdProps.bMaxChars = True
        End If 'stack
        If Not pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPoint And _
              pMpxOpLProps.CanOverrunFeature Then
            If pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon And _
                pMpxOpLProps.AllowAsymmetricOverrun Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Asymmetric Overrun")
                mxdProps.bAsymmetric = True
            Else
                sTmp = InsertTabs(lTabLevel) & "Overrun: " & pMpxOpLProps.MaximumLabelOverrun
                If m_Version >= 93 Then
                    sTmp = sTmp & " " & GetMaplexUnits((pMpxOpLProps2.MaximumLabelOverrunUnit))
                Else
                    sTmp = sTmp & " Points"
                End If
                sw.WriteLine(sTmp)
                mxdProps.bOverrun = True
            End If
        End If 'overrun
        If pMpxOpLProps.CanReduceFontSize Then
            If pTextSym.Size > pMpxOpLProps.FontHeightReductionLimit Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Font Reduction from " & pTextSym.Size & " to " & _
                             pMpxOpLProps.FontHeightReductionLimit & " step " & pMpxOpLProps.FontHeightReductionStep)
                mxdProps.bFontReduction = True
            End If
            If pMpxOpLProps.FontWidthReductionLimit < 100 Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Font Compression to " & pMpxOpLProps.FontWidthReductionLimit & _
                             " step " & pMpxOpLProps.FontWidthReductionStep)
                mxdProps.bCompression = True
            End If
        End If
        If pMpxOpLProps.CanAbbreviateLabel Then
            'careful here - if reading a layer file, pDictionaries has not been initialised
            If pMpxOpLProps.DictionaryName <> "" And Not mxdProps.pDictionaries Is Nothing Then
                If mxdProps.pDictionaries.DictionaryCount Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Abbreviation dictionary: " & pMpxOpLProps.DictionaryName)
                    mxdProps.bAbbreviation = True
                End If
            End If
            If pMpxOpLProps.CanTruncateLabel Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Truncation")
                mxdProps.bTruncation = True
                If m_Version >= 101 Then
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Min word length: " & pMpxOpLProps4.TruncationMinimumLength)
                    If pMpxOpLProps4.TruncationMinimumLength <> 1 Then mxdProps.bTruncationLength = True
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Marker character: " & pMpxOpLProps4.TruncationMarkerCharacter)
                    If pMpxOpLProps4.TruncationMarkerCharacter <> "." Then mxdProps.bTruncationMarker = True
                    mxdProps.sTruncMarker(mxdProps.lTruncMarker) = pMpxOpLProps4.TruncationMarkerCharacter
                    AddIfUnique(mxdProps.lTruncMarker, mxdProps.sTruncMarker, ARRAY_SIZE)
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Remove characters: " & pMpxOpLProps4.TruncationPreferredCharacters)
                    If StrComp(pMpxOpLProps4.TruncationPreferredCharacters, "aeiou", vbTextCompare) <> 0 Then mxdProps.bTruncationChars = True
                    mxdProps.sTruncChars(mxdProps.lTruncChars) = pMpxOpLProps4.TruncationPreferredCharacters
                    AddIfUnique(mxdProps.lTruncChars, mxdProps.sTruncChars, ARRAY_SIZE)
                End If '10.1
            End If
        End If
        If pMpxOpLProps.MinimumSizeForLabeling Then
            sTmp = InsertTabs(lTabLevel) & "Min size: " & pMpxOpLProps.MinimumSizeForLabeling
            If m_Version >= 93 Then
                sTmp = sTmp & " " & GetMaplexUnits((pMpxOpLProps2.MinimumFeatureSizeUnit))
            Else
                sTmp = sTmp & " Map Units"
            End If
            sw.WriteLine(sTmp)
            If m_Version >= 93 Then If pMpxOpLProps2.IsMinimumSizeBasedOnArea Then _
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Based on area")
            mxdProps.bMinSize = True
        End If
        If m_Version >= 93 Then
            mxdProps.sStrategyPriority(pMpxOpLProps2.StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyAbbreviation)) = "Abbreviate "
            mxdProps.sStrategyPriority(pMpxOpLProps2.StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontCompression)) = "Compress "
            mxdProps.sStrategyPriority(pMpxOpLProps2.StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontReduction)) = "Reduce "
            mxdProps.sStrategyPriority(pMpxOpLProps2.StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyOverrun)) = "Overrun "
            mxdProps.sStrategyPriority(pMpxOpLProps2.StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyStacking)) = "Stack "
            sTmp = "Strategy priority order: " & mxdProps.sStrategyPriority(1) & mxdProps.sStrategyPriority(2) & _
                         mxdProps.sStrategyPriority(3) & mxdProps.sStrategyPriority(4) & mxdProps.sStrategyPriority(5)
            'no overrun for points
            If pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPoint Then
                sTmp = Replace(sTmp, "Overrun ", "")
                sw.WriteLine(InsertTabs(lTabLevel) & sTmp)
                If StrComp(sTmp, "Strategy priority order: Stack Compress Reduce Abbreviate ") <> 0 Then mxdProps.bStrategyPriority = True
            Else
                sw.WriteLine(InsertTabs(lTabLevel) & sTmp)
                If StrComp(sTmp, "Strategy priority order: Stack Overrun Compress Reduce Abbreviate ") <> 0 Then mxdProps.bStrategyPriority = True
            End If
        End If '>=9.3
        If m_Version >= 101 Then
            If pMpxOpLProps4.CanKeyNumberLabel Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Key numbering")
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Group name: " & pMpxOpLProps4.KeyNumberGroupName)
                mxdProps.bKeyNumbering = True
            End If
        End If '10.1

        'Conflict tab
        sw.WriteLine(InsertTabs(lTabLevel - 1) & "Conflicts:")
        If pMpxOpLProps.FeatureWeight Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Weight: " & pMpxOpLProps.FeatureWeight)
            mxdProps.bWeights = True
        End If
        If (pMpxOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon And _
              pMpxOpLProps.PolygonBoundaryWeight) Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Boundary weight: " & pMpxOpLProps.PolygonBoundaryWeight)
            mxdProps.bWeights = True
        End If
        If pMpxOpLProps.BackgroundLabel Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Background label")
            mxdProps.bBackground = True
        End If
        If pMpxOpLProps.CanRemoveOverlappingLabel Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Can remove overlapping label (not used)")
        End If
        If pMpxOpLProps.ThinDuplicateLabels Then
            sTmp = InsertTabs(lTabLevel) & "Remove duplicates within: " & pMpxOpLProps.ThinningDistance
            If m_Version >= 93 Then
                sTmp = sTmp & " " & GetMaplexUnits((pMpxOpLProps2.ThinningDistanceUnit))
            Else
                sTmp = sTmp & " Map Units"
            End If
            sw.WriteLine(sTmp)
            mxdProps.bRemoveDup = True
        End If
        sw.WriteLine(InsertTabs(lTabLevel) & "Label buffer: " & pMpxOpLProps.LabelBuffer & "%")
        If pMpxOpLProps.LabelBuffer <> 15 Then mxdProps.bLabelBuffer = True
        If m_Version >= 93 Then
            If pMpxOpLProps2.IsLabelBufferHardConstraint Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Hard constraint")
                mxdProps.bHardConstraint = True
            End If
        End If '>=9.3
        If pMpxOpLProps.NeverRemoveLabel Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Never remove")
            mxdProps.bNeverRemove = True
        End If
        If pMpxOpLProps.FeatureBuffer Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Feature buffer: " & pMpxOpLProps.FeatureBuffer & " (not used)")
        End If

    End Sub

    Sub GetSLEProps(ByRef pLabelEngineLayerProperties As ILabelEngineLayerProperties2, ByRef lTabLevel As Integer)
        sw.Flush()
        Dim i As Integer
        Dim pBasicOpLProps As IBasicOverposterLayerProperties4
        Dim pOpLProps As IOverposterLayerProperties2
        'layer file - don't know if MLE or SLE
        Try
            pBasicOpLProps = pLabelEngineLayerProperties.BasicOverposterLayerProperties
            pOpLProps = pLabelEngineLayerProperties.OverposterLayerProperties
        Catch ex As Exception
            Return
        End Try

        sw.WriteLine(InsertTabs(lTabLevel - 1) & "Label properties:")
        'placement tab
        Dim dAngles() As Double
        If pBasicOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPoint Then
            'points
            If pBasicOpLProps.PointPlacementMethod = esriOverposterPointPlacementMethod.esriAroundPoint Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Place around point")
                sw.WriteLine(InsertTabs(lTabLevel + 1) & "Priorities (clockwise from north): " & _
                    pBasicOpLProps.PointPlacementPriorities.AboveCenter & " " & pBasicOpLProps.PointPlacementPriorities.AboveRight & " " & _
                    pBasicOpLProps.PointPlacementPriorities.CenterRight & " " & pBasicOpLProps.PointPlacementPriorities.BelowRight & " " & _
                    pBasicOpLProps.PointPlacementPriorities.BelowCenter & " " & pBasicOpLProps.PointPlacementPriorities.BelowLeft & " " & _
                    pBasicOpLProps.PointPlacementPriorities.CenterLeft & " " & pBasicOpLProps.PointPlacementPriorities.AboveLeft)
                mxdProps.bPointAround = True
            ElseIf pBasicOpLProps.PointPlacementMethod = esriOverposterPointPlacementMethod.esriOnTopPoint Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Place on top of point")
                mxdProps.bPointOnTop = True
            ElseIf pBasicOpLProps.PointPlacementMethod = esriOverposterPointPlacementMethod.esriRotationField Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Angle by field: " & pBasicOpLProps.RotationField)
                If pBasicOpLProps.RotationType = esriLabelRotationType.esriRotateLabelGeographic Then _
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Geographic style")
                If pBasicOpLProps.RotationType = esriLabelRotationType.esriRotateLabelArithmetic Then _
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Arithmetic style")
                If pBasicOpLProps.PerpendicularToAngle Then sw.WriteLine(InsertTabs(lTabLevel + 1) & "Place perpendicular")
                mxdProps.bPointRotation = True
            ElseIf pBasicOpLProps.PointPlacementMethod = esriOverposterPointPlacementMethod.esriSpecifiedAngles Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Place at specified angles:")
                dAngles = pBasicOpLProps.PointPlacementAngles()
                For i = 0 To UBound(dAngles)
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & dAngles(i))
                Next
                mxdProps.bPointSpecAngle = True
            End If 'point placement method
            'not needed?
            'If pBasicOpLProps.PointPlacementOnTop Then _
            'Print # InsertTabs(ltablevel) & "Place on top of point = true"
        ElseIf pBasicOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolyline Then
            'lines
            sw.WriteLine(InsertTabs(lTabLevel) & "Orientation:")
            If pBasicOpLProps.LineLabelPosition.Horizontal Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal")
                mxdProps.bLineHor = True
            End If
            If pBasicOpLProps.LineLabelPosition.Parallel Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Parallel")
                mxdProps.bLineParallel = True
            End If
            If pBasicOpLProps.LineLabelPosition.ProduceCurvedLabels Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Curved")
                mxdProps.bLineCrv = True
            End If
            If pBasicOpLProps.LineLabelPosition.Perpendicular Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Perpendicular")
                mxdProps.bLinePerp = True
            End If
            If Not pBasicOpLProps.LineLabelPosition.Horizontal Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Position:")
                If pBasicOpLProps.LineLabelPosition.Above Then sw.WriteLine(InsertTabs(lTabLevel) & "Above")
                If pBasicOpLProps.LineLabelPosition.OnTop Then sw.WriteLine(InsertTabs(lTabLevel) & "On top")
                If pBasicOpLProps.LineLabelPosition.Below Then sw.WriteLine(InsertTabs(lTabLevel) & "Below")
                If pBasicOpLProps.LineLabelPosition.Left Then sw.WriteLine(InsertTabs(lTabLevel) & "Left")
                If pBasicOpLProps.LineLabelPosition.Right Then sw.WriteLine(InsertTabs(lTabLevel) & "Right")
                If pBasicOpLProps.LineOffset Then
                    sw.WriteLine(InsertTabs(lTabLevel) & "Offset: " & pBasicOpLProps.LineOffset)
                    mxdProps.bLineOffset = True
                End If
            End If 'not horizontal
            If pBasicOpLProps.LineLabelPosition.Parallel Or pBasicOpLProps.LineLabelPosition.Perpendicular Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Location:")
                If pBasicOpLProps.LineLabelPosition.InLine Then sw.WriteLine(InsertTabs(lTabLevel) & "Best")
                If pBasicOpLProps.LineLabelPosition.AtEnd Then sw.WriteLine(InsertTabs(lTabLevel) & "At end")
                If pBasicOpLProps.LineLabelPosition.AtStart Then sw.WriteLine(InsertTabs(lTabLevel) & "At start")
                If pBasicOpLProps.LineLabelPosition.AtEnd Or pBasicOpLProps.LineLabelPosition.AtStart Then
                    sw.WriteLine(InsertTabs(lTabLevel + 1) & "Non-zero priorities:")
                    If pBasicOpLProps.LineLabelPlacementPriorities.AboveAfter Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Above after = " & pBasicOpLProps.LineLabelPlacementPriorities.AboveAfter)
                    If pBasicOpLProps.LineLabelPlacementPriorities.AboveAlong Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Above along = " & pBasicOpLProps.LineLabelPlacementPriorities.AboveAlong)
                    If pBasicOpLProps.LineLabelPlacementPriorities.AboveBefore Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Above before = " & pBasicOpLProps.LineLabelPlacementPriorities.AboveBefore)
                    If pBasicOpLProps.LineLabelPlacementPriorities.AboveEnd Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Above end = " & pBasicOpLProps.LineLabelPlacementPriorities.AboveEnd)
                    If pBasicOpLProps.LineLabelPlacementPriorities.AboveStart Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Above start = " & pBasicOpLProps.LineLabelPlacementPriorities.AboveStart)
                    If pBasicOpLProps.LineLabelPlacementPriorities.BelowAfter Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Below after = " & pBasicOpLProps.LineLabelPlacementPriorities.BelowAfter)
                    If pBasicOpLProps.LineLabelPlacementPriorities.BelowAlong Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Below along = " & pBasicOpLProps.LineLabelPlacementPriorities.BelowAlong)
                    If pBasicOpLProps.LineLabelPlacementPriorities.BelowBefore Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Below before = " & pBasicOpLProps.LineLabelPlacementPriorities.BelowBefore)
                    If pBasicOpLProps.LineLabelPlacementPriorities.BelowEnd Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Below end = " & pBasicOpLProps.LineLabelPlacementPriorities.BelowEnd)
                    If pBasicOpLProps.LineLabelPlacementPriorities.BelowStart Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Below start = " & pBasicOpLProps.LineLabelPlacementPriorities.BelowStart)
                    If pBasicOpLProps.LineLabelPlacementPriorities.CenterAfter Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Center after = " & pBasicOpLProps.LineLabelPlacementPriorities.CenterAfter)
                    If pBasicOpLProps.LineLabelPlacementPriorities.CenterAlong Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Center along = " & pBasicOpLProps.LineLabelPlacementPriorities.CenterAlong)
                    If pBasicOpLProps.LineLabelPlacementPriorities.CenterBefore Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Center before = " & pBasicOpLProps.LineLabelPlacementPriorities.CenterBefore)
                    If pBasicOpLProps.LineLabelPlacementPriorities.CenterEnd Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Center end = " & pBasicOpLProps.LineLabelPlacementPriorities.CenterEnd)
                    If pBasicOpLProps.LineLabelPlacementPriorities.CenterStart Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Center start = " & pBasicOpLProps.LineLabelPlacementPriorities.CenterStart)
                    If pBasicOpLProps.LineLabelPosition.Offset Then _
                        sw.WriteLine(InsertTabs(lTabLevel + 2) & "Offset = " & pBasicOpLProps.LineLabelPosition.Offset)
                End If
            End If
            If pBasicOpLProps.MaxDistanceFromTarget Then sw.WriteLine("Max distance from target: " & pBasicOpLProps.MaxDistanceFromTarget)
        ElseIf pBasicOpLProps.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon Then
            'polygons
            If pBasicOpLProps.PolygonPlacementMethod = esriOverposterPolygonPlacementMethod.esriAlwaysHorizontal Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Horizontal")
                mxdProps.bPolyHorz = True
            End If
            If pBasicOpLProps.PolygonPlacementMethod = esriOverposterPolygonPlacementMethod.esriAlwaysStraight Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Straight")
                mxdProps.bPolyStr = True
            End If
            If pBasicOpLProps.PolygonPlacementMethod = esriOverposterPolygonPlacementMethod.esriMixedStrategy Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Try horizontal first")
                mxdProps.bPolyTryHorz = True
            End If
            If pBasicOpLProps.PlaceOnlyInsidePolygon Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Only place inside")
                mxdProps.bPolyPlaceOnlyInside = True
            End If
        End If 'feature type
        If pBasicOpLProps.FeatureType <> esriBasicOverposterFeatureType.esriOverposterPoint Then
            If pBasicOpLProps.NumLabelsOption = esriBasicNumLabelsOption.esriNoLabelRestrictions Then
                sw.WriteLine(InsertTabs(lTabLevel) & "No Label Restrictions")
                mxdProps.bNumLabNoRestrict = True
            End If
            If pBasicOpLProps.NumLabelsOption = esriBasicNumLabelsOption.esriOneLabelPerShape Then
                sw.WriteLine(InsertTabs(lTabLevel) & "One label per feature")
                mxdProps.bNumLabperName = True
            End If
            If pBasicOpLProps.NumLabelsOption = esriBasicNumLabelsOption.esriOneLabelPerPart Then
                sw.WriteLine(InsertTabs(lTabLevel) & "One label per feature part")
                mxdProps.bNumLabperPart = True
            End If
            If pBasicOpLProps.NumLabelsOption = esriBasicNumLabelsOption.esriOneLabelPerName Then
                sw.WriteLine(InsertTabs(lTabLevel) & "Remove duplicate labels")
                mxdProps.bRemoveDupSLE = True
            End If
        End If 'not point

        'conflict tab
        sw.WriteLine(InsertTabs(lTabLevel - 1) & "Conflicts:")
        If pBasicOpLProps.FeatureWeight Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Feature weight: " & pBasicOpLProps.FeatureWeight)
            mxdProps.bWeights = True
        End If
        If pBasicOpLProps.LabelWeight Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Label weight: " & pBasicOpLProps.LabelWeight)
            If pBasicOpLProps.LabelWeight <> esriBasicOverposterWeight.esriHighWeight Then mxdProps.bWeights = True
        End If
        If pBasicOpLProps.BufferRatio Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Buffer ratio: " & pBasicOpLProps.BufferRatio)
            mxdProps.bLabelBuffer = True
        End If
        If Not pOpLProps.TagUnplaced Then
            sw.WriteLine(InsertTabs(lTabLevel) & "Allow overlapping labels")
            mxdProps.bOverlappingLabels = True
        End If

    End Sub

    Function GetUnits(ByRef lUnits As Integer) As String
        Select Case lUnits
            Case 1
                GetUnits = "Inches"
            Case 2
                GetUnits = "Points"
            Case 3
                GetUnits = "Feet"
            Case 4
                GetUnits = "Yards"
            Case 5
                GetUnits = "Miles"
            Case 6
                GetUnits = "Nautical miles"
            Case 7
                GetUnits = "Millimeters"
            Case 8
                GetUnits = "Centimeters"
            Case 9
                GetUnits = "Meters"
            Case 10
                GetUnits = "Kilometers"
            Case 11
                GetUnits = "Decimal degrees"
            Case 12
                GetUnits = "Decimeters"
            Case Else
                GetUnits = "Unknown"
        End Select
    End Function

    Function GetMaplexUnits(ByRef lUnits As Integer) As String
        Select Case lUnits
            Case 0
                GetMaplexUnits = "Map Units"
            Case 1
                GetMaplexUnits = "Millimeters"
            Case 2
                GetMaplexUnits = "Inches"
            Case 3
                GetMaplexUnits = "Points"
            Case 4
                GetMaplexUnits = "%"
            Case Else
                GetMaplexUnits = "Unknown units"
        End Select
    End Function

    Function GetCharSet(ByRef lCharSet As Integer) As String
        Select Case (lCharSet)
            Case &H0
                GetCharSet = "ANSI Charset"
            Case &H1
                GetCharSet = "Default Charset"
            Case &H2
                GetCharSet = "Symbol Charset"
            Case &H4D
                GetCharSet = "Apple Mac Charset"
            Case &H80
                GetCharSet = "Japanese Charset"
            Case &H81
                GetCharSet = "Hangul Charset"
            Case &H82
                GetCharSet = "Johab Charset"
            Case &H86
                GetCharSet = "Simplified Chinese Charset"
            Case &H88
                GetCharSet = "Traditional Chinese Charset"
            Case &HA1
                GetCharSet = "Greek Charset"
            Case &HA2
                GetCharSet = "Turkish Charset"
            Case &HA3
                GetCharSet = "Vietnamese Charset"
            Case &HB1
                GetCharSet = "Hebrew Charset"
            Case &HB2
                GetCharSet = "Arabic Charset"
            Case &HBA
                GetCharSet = "Baltic Charset"
            Case &HCC
                GetCharSet = "Russian Charset"
            Case &HDE
                GetCharSet = "Thai Charset"
            Case &HEE
                GetCharSet = "Eastern Europe Charset"
            Case &HFF
                GetCharSet = "OEM Charset"
            Case Else
                GetCharSet = "Unknown Charset"
        End Select
    End Function

    Function GetGeomType(ByRef lGeom As Integer) As String
        Select Case (lGeom)
            Case 0
                GetGeomType = "Null"
            Case 1
                GetGeomType = "Point"
            Case 2
                GetGeomType = "Multipoint"
            Case 3
                GetGeomType = "Polyline"
            Case 4
                GetGeomType = "Polygon"
            Case 5
                GetGeomType = "Envelope"
            Case 6
                GetGeomType = "Path"
            Case 7
                GetGeomType = "Any"
            Case 9
                GetGeomType = "Multipatch"
            Case 13
                GetGeomType = "Line"
            Case 17
                GetGeomType = "Geometry Bag"
            Case Else
                GetGeomType = "Other (" & CStr(lGeom) & ")"
        End Select
    End Function

    Function GetRGB(ByRef pColor As IColor) As String
        GetRGB = "No color"
        If pColor Is Nothing Then Exit Function

        Dim pRGB As IRgbColor
        Dim a, l, b As Double

        pColor.GetCIELAB(l, a, b)
        pRGB = New RgbColor
        pRGB.SetCIELAB(l, a, b)
        GetRGB = CStr(pRGB.Red) & ", " & CStr(pRGB.Green) & ", " & CStr(pRGB.Blue)
    End Function

    Function GetCMYK(ByRef pColor As IColor) As String
        GetCMYK = "No color"
        If pColor Is Nothing Then Exit Function

        Dim pCMYK As ICmykColor
        Dim a, l, b As Double

        pColor.GetCIELAB(l, a, b)
        pCMYK = New CmykColor
        pCMYK.SetCIELAB(l, a, b)
        GetCMYK = CStr(pCMYK.Cyan) & ", " & CStr(pCMYK.Magenta) & ", " & CStr(pCMYK.Yellow) & ", " & CStr(pCMYK.Black)
    End Function

    Function InsertTabs(ByRef lTabLevel As Integer) As String
        If lTabLevel < 1 Then
            InsertTabs = ""
        Else
            InsertTabs = InsertTabs(lTabLevel - 1) & vbTab
        End If
    End Function

    Function GetClipType(ByRef lType As Integer) As String
        Select Case (lType)
            Case 0
                GetClipType = "None"
            Case 1
                GetClipType = "Shape"
            Case 2
                GetClipType = "Extent"
            Case 3
                GetClipType = "Page Index"
            Case Else
                GetClipType = "Unknown Type"
        End Select
    End Function

    Sub GetExtentInfo(ByRef eType As esriExtentTypeEnum, ByRef eBounds As IEnvelope, ByRef dScale As Double)
        Select Case eType
            Case esriExtentTypeEnum.esriExtentBounds
                sw.WriteLine(InsertTabs(1) & "Map Frame Fixed Extent")
                If eBounds.IsEmpty() Then
                    sw.WriteLine(InsertTabs(2) & "Error: Bounds empty")
                Else
                    sw.WriteLine(InsertTabs(2) & "Bounds: " & eBounds.XMin & ", " & eBounds.YMin & ", " & eBounds.XMax & ", " & eBounds.YMax)
                End If
                mxdProps.bFixedExtent = True
            Case esriExtentTypeEnum.esriExtentScale
                sw.WriteLine(InsertTabs(1) & "Map Frame Fixed Scale")
                sw.WriteLine(InsertTabs(2) & "Scale: " & dScale)
                mxdProps.bFixedScale = True
            Case esriExtentTypeEnum.esriExtentDefault
                sw.WriteLine(InsertTabs(1) & "Map Frame Automatic Extent")
                mxdProps.bAutoExtent = True
            Case esriExtentTypeEnum.esriAutoExtentMarginPercent
                sw.WriteLine(InsertTabs(1) & "Map Frame Margin in Percent")
            Case esriExtentTypeEnum.esriAutoExtentMarginMapUnits
                sw.WriteLine(InsertTabs(1) & "Map Frame Margin in Map Units")
            Case esriExtentTypeEnum.esriAutoExtentMarginPageUnits
                sw.WriteLine(InsertTabs(1) & "Map Frame Margin in Page Units")
                'Case esriExtentTypeEnum.esriAutoExtentFeatures
                '  sw.WriteLine( InsertTabs(1) & "Map frame extent intersected with features in another data frame")
            Case esriExtentTypeEnum.esriExtentPageIndex
                sw.WriteLine(InsertTabs(1) & "Map Frame Page Index Extent")
            Case Else
                sw.WriteLine(InsertTabs(1) & "Map Frame Extent Type Unknown - " & CShort(eType))
        End Select
    End Sub

    ' compare last item in array to the other values,
    ' if it is unique, increment the counter (i.e. keep it).
    ' easier to do it this way because of the way QuerySeparators works
    Public Sub AddIfUnique(ByRef counter As Long, ByRef sArray() As String, ByVal lMax As Long)
        Dim j As Long
        'flag if array full
        For j = 0 To counter
            If j = counter And counter < lMax Then
                counter = counter + 1
            Else
                If StrComp(sArray(j), sArray(counter)) = 0 Then Exit For
            End If
        Next

    End Sub

    'count selected features (could extend in the future)
    'From ArcGIS Resource Center: How to zoom to selected features in globe
    'http://help.arcgis.com/en/sdk/10.0/arcobjects_net/conceptualhelp/index.html#/d/0001000000vn000000.htm
    '
    Function GetSelectedFeatures(ByRef pLayer As ILayer, ByRef lTabLevel As Integer) As Integer
        'Dim featureLayer As ESRI.ArcGIS.Carto.IFeatureLayer = CType(pLayer, ESRI.ArcGIS.Carto.IFeatureLayer) ' Explicit Cast
        Dim featureSelection As ESRI.ArcGIS.Carto.IFeatureSelection = CType(pLayer, ESRI.ArcGIS.Carto.IFeatureSelection) ' Explicit Cast
        Dim selectionSet As ESRI.ArcGIS.Geodatabase.ISelectionSet = featureSelection.SelectionSet

        GetSelectedFeatures = 0
        '    Dim featureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = featureLayer.FeatureClass

        '    Dim shapeField As System.String = featureClass.ShapeFieldName
        '    spatialFilterCls.GeometryField = shapeField
        '    spatialReference = spatialFilterCls.OutputSpatialReference(shapeField)

        '    Dim cursor As ESRI.ArcGIS.Geodatabase.ICursor = Nothing
        '    selectionSet.Search(spatialFilterCls, True, cursor)
        '    Dim featureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor = CType(cursor, ESRI.ArcGIS.Geodatabase.IFeatureCursor) ' Explicit Cast

        '    Dim getLayerExtent As System.Boolean = True
        '    Dim feature As ESRI.ArcGIS.Geodatabase.IFeature

        '    feature = featureCursor.NextFeature

        '    While (feature) IsNot Nothing
        '      Dim geometry As ESRI.ArcGIS.Geometry.IGeometry = feature.Shape
        '      Dim featureExtent As ESRI.ArcGIS.Geometry.IEnvelope = geometry.Envelope
        '      envelopeCls.Union(featureExtent)

        '      haveFeatures = True

        '      If getLayerExtent Then
        '        Dim geoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset = CType(featureLayer, ESRI.ArcGIS.Geodatabase.IGeoDataset) ' Explicit Cast
        '        If Not (geoDataset Is Nothing) Then
        '          Dim layerExtent As ESRI.ArcGIS.Geometry.IEnvelope = geoDataset.Extent
        '          layersExtentCls.Union(layerExtent)
        '        End If
        '        getLayerExtent = False
        '      End If

        '      feature = featureCursor.NextFeature ' Iterate through the next feature
        '    End While
        'End If ' typeof
        '    layer = enumLayer.Next() ' Iterate through the next layer
        'End While
        If Not selectionSet Is Nothing Then GetSelectedFeatures = selectionSet.Count

    End Function

    Function GetPointPlacementMethod(ByRef lMethod As Integer) As String
        Select Case lMethod
            Case 0
                GetPointPlacementMethod = "Best position"
            Case 1
                GetPointPlacementMethod = "Centered"
            Case 2
                GetPointPlacementMethod = "North"
            Case 3
                GetPointPlacementMethod = "Northeast"
            Case 4
                GetPointPlacementMethod = "East"
            Case 5
                GetPointPlacementMethod = "Southeast"
            Case 6
                GetPointPlacementMethod = "South"
            Case 7
                GetPointPlacementMethod = "Southwest"
            Case 8
                GetPointPlacementMethod = "West"
            Case 9
                GetPointPlacementMethod = "Northwest"
            Case Else
                GetPointPlacementMethod = "Unknown position"
        End Select
    End Function

    Function GetOverposterWeight(ByRef lWeight As Integer) As String
        Select Case (lWeight)
            Case 0
                GetOverposterWeight = "No Weight"
            Case 1
                GetOverposterWeight = "Low Weight"
            Case 2
                GetOverposterWeight = "Medium Weight"
            Case 3
                GetOverposterWeight = "High Weight"
            Case Else
                GetOverposterWeight = "Unknown Weight"
        End Select
    End Function

    Function GetSymbolSubstitutionType(ByRef lType As Integer) As String
        Select Case (lType)
            Case 0
                GetSymbolSubstitutionType = "None"
            Case 1
                GetSymbolSubstitutionType = "Color"
            Case 2
                GetSymbolSubstitutionType = "Individual Subordinate"
            Case 3
                GetSymbolSubstitutionType = "Individual Dominant"
            Case Else
                GetSymbolSubstitutionType = "Unknown Type"
        End Select
    End Function

    Function GetPicType(ByRef lType As Integer) As String
        Select Case (lType)
            Case -1
                GetPicType = "Unitialized"
            Case 0
                GetPicType = "None"
            Case 1
                GetPicType = "Bitmap"
            Case 2
                GetPicType = "Metafile"
            Case 3
                GetPicType = "Icon"
            Case 4
                GetPicType = "Enhanced Metafile"
            Case Else
                GetPicType = "Unknown Type"
        End Select
    End Function

    Function GetLineCalloutStyle(ByVal style As esriLineCalloutStyle) As String
        Select Case (style)
            Case esriLineCalloutStyle.esriLCSBase
                GetLineCalloutStyle = "Base"
            Case esriLineCalloutStyle.esriLCSCircularCCW
                GetLineCalloutStyle = "Circular Counterclockwise"
            Case esriLineCalloutStyle.esriLCSCircularCW
                GetLineCalloutStyle = "Circular Clockwise"
            Case esriLineCalloutStyle.esriLCSCustom
                GetLineCalloutStyle = "Custom"
            Case esriLineCalloutStyle.esriLCSFourPoint
                GetLineCalloutStyle = "Four Point"
            Case esriLineCalloutStyle.esriLCSMidpoint
                GetLineCalloutStyle = "Midpoint"
            Case esriLineCalloutStyle.esriLCSThreePoint
                GetLineCalloutStyle = "Three point"
            Case esriLineCalloutStyle.esriLCSUnderline
                GetLineCalloutStyle = "Underline"
            Case Else
                GetLineCalloutStyle = "Unknown style"
        End Select
    End Function

    Function GetColorRampAlgorithm(ByVal algorithm As esriColorRampAlgorithm) As String
        Select Case (algorithm)
            Case esriColorRampAlgorithm.esriCIELabAlgorithm
                GetColorRampAlgorithm = "CIELab"
            Case esriColorRampAlgorithm.esriHSVAlgorithm
                GetColorRampAlgorithm = "HSV"
            Case esriColorRampAlgorithm.esriLabLChAlgorithm
                GetColorRampAlgorithm = "LabLCh"
            Case Else
                GetColorRampAlgorithm = "Unknown algorithm"
        End Select
    End Function

    'is name qualified with a possible join table name?
    Function IsQualifiedName(ByVal sName As String) As Boolean

        IsQualifiedName = False

        'if no full stops, then not qualified
        If InStr(sName, ".") = 0 Then Return False

        'if no spaces then not an expression, return true
        If InStr(sName, " ") = 0 Then
            mxdProps.bQualifiedNames = True
            Return True
        End If

        'expression - look for full stop between square brackets
        Dim m As Match = Regex.Match(sName, "\[\w*\.\w*\]")
        If m.Success Then
            mxdProps.bQualifiedNames = True
            IsQualifiedName = True
        End If

    End Function

    ' Find possible coded value domains for a given field of a feature layer
    Sub GetCodedValueDomain(ByRef pGeoFL As IGeoFeatureLayer, ByVal sFieldName As String, ByVal lTabLevel As Long)

        Dim featureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pGeoFL.FeatureClass
        If featureClass Is Nothing Then Return

        Dim fields As IFields = featureClass.Fields
        Dim fieldIndex As Integer = featureClass.FindField(sFieldName)
        If fieldIndex <> -1 Then
            ' Found field, see if there is a domain
            Dim field As IField = fields.Field(fieldIndex)
            If Not field.Domain Is Nothing Then
                If field.Domain.Type = esriDomainType.esriDTCodedValue Then
                    sw.WriteLine(InsertTabs(lTabLevel + 2) & "Coded value domain for [{0}]: {1}", sFieldName, field.Domain.Name)
                    mxdProps.bCodedValueDomain = True
                End If
            End If
        End If

    End Sub
End Module