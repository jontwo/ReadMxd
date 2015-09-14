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

Imports ESRI.ArcGIS.Carto


Public Class clsMxdProps
    'MLE
    'Points
    Public bMayShift, bPointFixed, bPointBest, bPointZones, bAlteredZones As Boolean
    Public bPointMaxOffset, bPointOffDist, bPointFtrGeom As Boolean
    Public bPointGAStr, bPointGACurv, bPointGACurvNoFlip, bPointGAStrNoFlip As Boolean
    Public bPointRotAngle, bPointRotation, bPointRotFlip As Boolean
    'Lines
    Public bLineCenCur, bLineCenHor, bLineCenStr, bLineCenPer As Boolean
    Public bLineOffCur, bLineOffHor, bLineOffStr, bLineOffPer As Boolean
    Public bLineStreetAdd, bLineRegular, bLineStreet, bLineContour As Boolean
    Public bLineOffDist, bLineRiver As Boolean ', bLinePrefOffset
    Public bLineBeforeStart, bLineBestAlong, bLineAfterEnd As Boolean
    Public bLineFromStart, bLineFromEnd, bLineFtrGeom As Boolean
    Public bNoConstraint, bConstrainAbove, bConstrainBelow As Boolean
    Public bConstrainLeft, bConstrainRight As Boolean
    Public bStreetPrimary, bStreetHorz, bStreetReduce, bStreetSpread As Boolean
    Public bContourPage, bLineSecOff, bContourUphill, bContourLadder As Boolean
    Public bLineRepeat, bContourNoLadder, bLineDirection, bLineSpread As Boolean
    Public bLineGAStr, bLineGACurv, bLineGACurvNoFlip, bLineGAStrNoFlip As Boolean
    'Polygons
    Public bPolyOffHorz, bPolyStr, bPolyHorz, bPolyCurv, bPolyBdy As Boolean
    Public bPolyRiver, bPolyRegular, bPolyOffCurv, bPolyParcel, bPolyTryHorz As Boolean
    Public bPolyExtZones, bPolyMayPlaceOutside, bPolyIntZones, bPolyAnchor As Boolean
    Public bPolySpread, bPolyRepeat, bPolyAllowHoles As Boolean
    Public bPolyBdyAllowHoles, bPolyBdySingleSided, bPolyBdyOnLine As Boolean
    Public bPolyMaxOffset, bPolyOffDist, bPolyFtrGeom As Boolean
    Public bPolyGAStr, bPolyGACurv, bPolyGACurvNoFlip, bPolyGAStrNoFlip As Boolean
    'Strategies
    Public bAsymmetric, bStack, bOverrun, bFontReduction As Boolean
    Public bAbbreviation, bCompression, bTruncation, bMinSize As Boolean
    Public bStackL, bStackC, bStackLorR, bStackR As Boolean
    Public bSepFor, bSepVis, bSepAft As Boolean
    Public bMinChars, bMaxChars, bMaxLines As Boolean
    Public bStrategyPriority As Boolean
    Public sStrategyPriority(5) As String
    Public sSeparators(ARRAY_SIZE) As String
    Public lSeparators As Long
    Public sTruncMarker(ARRAY_SIZE), sTruncChars(ARRAY_SIZE) As String
    Public lTruncMarker, lTruncChars As Long
    'Conflicts
    Public bWeights, bRemoveDup, bBackground, bLabelBuffer, bHardConstraint, bNeverRemove As Boolean

    'SLE
    Public bLinePerp, bLineParallel, bLineHor, bLineCrv, bLineOffset, bPolyPlaceOnlyInside As Boolean
    Public bNumLabperPart, bNumLabNoRestrict, bNumLabperName, bRemoveDupSLE As Boolean
    Public bPointOnTop, bOverlappingLabels, bPointAround, bPointSpecAngle As Boolean

    'Misc
    Public bLayerDefQuery, bScaleRanges, bQualifiedNames, bHTMLEnt, bCodedValueDomain As Boolean
    Public bLabelExpression, bLabelPriority, bSQL, bUninitPriority, bTags As Boolean
    Public bTextCase, bRighttoLeft, bBaseTag, bTextPosition, bCharSpacing As Boolean
    Public bKerningOff, bCharWidth, bLeading, bWordSpacing, bXYOffset As Boolean
    Public bLineCallout, bTextBackground, bFillSymbol, bBalloonCallout, bMarkerTextBkg As Boolean
    Public bCJK, bSimpleLineCallout, bScaletoFit, bHalo, bShadow As Boolean
    Public bMultipoint, bFixedScale, bFixedExtent, bAutoExtent, bMultipatch As Boolean
    Public bSHP, bFGDB, bPGDB, bSDE, bCoverage As Boolean
    Public bMLE, bSLE, bMapIsMLE, bMapIsSLE As Boolean 'mxd contains some mle/sle props vs current map uses mle/sle

    '10.1
    Public bKeyNumbering, bKNDelimiter, bKNLeft, bKNRight, bKNAuto As Boolean
    Public bKNMax, bKNMin, bKNMayReset, bKNAlwaysReset, bKNNoReset As Boolean
    Public bSymbolOutline, bLabelNearJunction, bLabelNearBorder As Boolean
    Public bTruncationMarker, bTruncationLength, bTruncationChars As Boolean
    Public bMultiOptionFeature, bMultiOptionPart, bMultiOptionSegment As Boolean
    Public bStraddlacking, bWhitespace, bLinebreaks As Boolean

    'symbols
    Public bPieChart, bBarChart, bStackedChart, bSimpleFill, bGradientFill, bPictureFill, bMarkerFill, bLineFill As Boolean
    Public bFixedSize, bChartLeaders, b3DChart, bChartOverlap, bBarOrient, bColumnOrient, bGeogOrient, bArithOrient As Boolean
    Public bColorRamp, bRasterClassify, bRasterRGB, bRasterUnique, bRasterDiscrete, bRasterStretch As Boolean
    'TODO more symbols on summary

    'map props
    Public bGeographic, bRefScale, bProjected, bClipExtent As Boolean
    Public bFrameRotation, bClipToShape, bExcludeLayers, bLayoutView As Boolean
    Public bAllowOverlap, bMinimize, bUnambiguous, bLargestOnly As Boolean
    Public bDictionaryKeyword, bFast, bDictionaryTranslation, bDictionaryEnding As Boolean
    Public bDrawUnplaced, bRotateWithDataFrame As Boolean
    Public sInvertedLabTol As String = "" 'show tol for each dataframe
    Public lMapCount As Long = 1
    Public bRelPaths As Boolean, bAbsPaths As Boolean

    Public lAnnoLayers, lLabelClassCount, lAnnoLabelClassCount, lBarriers As Integer
    Public sDataSources(ARRAY_SIZE), sSRef(), sMapUnits() As String
    Public lDataSources As Long
    Public pDictionaries As IMaplexDictionaries
    Public pKeyNumberGroups As IMaplexKeyNumberGroups

End Class
