' (C) Copyright 2002-2005 by Autodesk, Inc. 
'
' Permission to use, copy, modify, and distribute this software in
' object code form for any purpose and without fee is hereby granted, 
' provided that the above copyright notice appears in all copies and 
' that both that copyright notice and the limited warranty and
' restricted rights notice below appear in all supporting 
' documentation.
'
' AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
' AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
' DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
' UNINTERRUPTED OR ERROR FREE.
'
' Use, duplication, or disclosure by the U.S. Government is subject to 
' restrictions set forth in FAR 52.227-19 (Commercial Computer
' Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
' (Rights in Technical Data and Computer Software), as applicable.
'
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Imports Excel = Microsoft.Office.Interop.Excel



Imports System
Imports System.IO
Imports System.Text
Imports System.Xml


Public Class Commands
  Dim tm As Autodesk.AutoCAD.DatabaseServices.TransactionManager
  Dim xmdoc As XmlDocument = New XmlDocument()
  Dim gsView As Autodesk.AutoCAD.GraphicsSystem.View
  Structure EstateDatas
    Dim Number As String
    Dim AddInfo As String
    Dim Street As String
    Dim ZipCode As String
    Dim Fid As Long
    Dim X As Double
    Dim Y As Double
  End Structure

  Structure AddressDatas
    Dim Number As String
    Dim AddInfo As String
    Dim Street As String
    Dim Quarters As Integer
    Dim Clients As Integer
    Dim PDAll As Integer
    Dim PDActive As Integer
    Dim TKT As Boolean
    Dim Gpon As Boolean
    Dim Fid As Long
  End Structure

  Dim Address() As EstateDatas
  Dim ad() As AddressDatas

  Public Function ExportAddress(ByVal ed As Editor) As XmlElement
    Dim xmlAddress As XmlElement = xmdoc.CreateElement("HouseNumbers")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCHOUSENUMBER"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт номеров домов.")
    ed.UpdateScreen()
    Res = ed.SelectAll(sfilter)
    If Res.Status = PromptStatus.Error Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("Чертеж не содержит номеров домов.")
      ed.UpdateScreen()
      Return xmlAddress
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    'Dim count As Integer
    'count = IdArray.Length
    Dim Id As ObjectId
    Dim c As Long = 0
    ReDim Address(SS.Count)
    For Each Id In IdArray
      Dim hn As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmHouseNumber As XmlElement = xmdoc.CreateElement("HouseNumber")
      xmHouseNumber.SetAttribute("handle", hn.Handle.ToString())
      Dim xmNumber As XmlElement = xmdoc.CreateElement("Number")
      Dim xmAddInfo As XmlElement = xmdoc.CreateElement("AddInfo")
      Dim xmStreet As XmlElement = xmdoc.CreateElement("StreetName")
      Dim xmOid As XmlElement = xmdoc.CreateElement("Oid")
      Dim xmZipCode As XmlElement = xmdoc.CreateElement("ZipCode")
      Dim xmPointX As XmlElement = xmdoc.CreateElement("X")
      Dim xmPointY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      xmPointX.InnerText = hn.Position.X.ToString()
      xmPointY.InnerText = hn.Position.Y.ToString()
      xmRotation.InnerText = hn.Rotation.ToString()
      Dim rs As Array = hn.GetXDataForApplication("ESTATE_DATAS").AsArray()
      xmStreet.InnerText = rs(1).Value
      Address(c).Street = rs(1).Value
      xmNumber.InnerText = rs(3).Value
      Address(c).Number = rs(3).Value
      xmAddInfo.InnerText = rs(4).Value
      Address(c).AddInfo = rs(4).Value
      xmZipCode.InnerText = rs(11).Value
      Address(c).ZipCode = rs(11).Value


      Dim rsFID As ResultBuffer = hn.GetXDataForApplication("FID")
      If rsFID = Nothing Then
        xmOid.InnerText = "0"
        Address(c).Fid = 0
      Else
        xmOid.InnerText = rsFID.AsArray(2).Value
        Address(c).Fid = rsFID.AsArray(2).Value
      End If

      xmHouseNumber.AppendChild(xmOid)
      xmHouseNumber.AppendChild(xmNumber)
      xmHouseNumber.AppendChild(xmAddInfo)
      xmHouseNumber.AppendChild(xmStreet)
      xmHouseNumber.AppendChild(xmZipCode)
      xmHouseNumber.AppendChild(xmPointX)
      xmHouseNumber.AppendChild(xmPointY)
      xmHouseNumber.AppendChild(xmRotation)
      xmlAddress.AppendChild(xmHouseNumber)
      c = c + 1
    Next


    Return xmlAddress
  End Function
  Public Function ExportMap(ByVal ed As Editor) As XmlElement
    Dim layer As XmlElement = xmdoc.CreateElement("Houses")
        Dim values() As TypedValue = {New TypedValue(DxfCode.LayerName, "HOUSE_GRID,GROUND_PLANE*,STRUCTURE_DOM"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт адресной информации.")
    ed.UpdateScreen()
    Res = ed.SelectAll(sfilter)
    If Res.Status = PromptStatus.Error Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("Чертеж не содержит адресной информации")
      ed.UpdateScreen()
      Return layer
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    'Dim count As Integer
    'count = IdArray.Length
    Try
      Dim Id As ObjectId


      For Each Id In IdArray
        Dim Ent As Entity = tm.GetObject(Id, OpenMode.ForRead, False)
        'ed.WriteMessage((ControlChars.Lf + "You selected: " + Ent.GetType().FullName))
        'Dim s As String = Ent.GetType().Name.ToString()
        Select Case Ent.GetType().Name.ToString()
          Case "Line"
            'Dim rowHouseLine As DataRow
            Dim ln As Line = tm.GetObject(Id, OpenMode.ForRead)
            Dim xline As XmlElement = xmdoc.CreateElement("Line")
            xline.SetAttribute("handle", ln.Handle.ToString())
            Dim xmpoint1 As XmlElement = xmdoc.CreateElement("Point")
            Dim xmpoint2 As XmlElement = xmdoc.CreateElement("Point")
            Dim StartPoint As StringBuilder = New StringBuilder()
            Dim EndPoint As StringBuilder = New StringBuilder()

            StartPoint.Append(ln.StartPoint.X)
            StartPoint.Append(", ")
            StartPoint.Append(ln.StartPoint.Y)
            xmpoint1.InnerText = StartPoint.ToString()
            xline.AppendChild(xmpoint1)

            EndPoint.Append(ln.EndPoint.X)
            EndPoint.Append(", ")
            EndPoint.Append(ln.EndPoint.Y)
            xmpoint2.InnerText = EndPoint.ToString()
            xline.AppendChild(xmpoint2)
            layer.AppendChild(xline)

          Case "Polyline"
            Dim pline As Polyline = tm.GetObject(Id, OpenMode.ForRead)
            Dim xpline As XmlElement = xmdoc.CreateElement("Polyline")
            xpline.SetAttribute("handle", pline.Handle.ToString())
            For c As Integer = 0 To pline.NumberOfVertices - 1
              Dim xmpoint As XmlElement = xmdoc.CreateElement("Point")
              Dim t As StringBuilder = New StringBuilder()
              t.Append(pline.GetPoint2dAt(c).X)
              t.Append(", ")
              t.Append(pline.GetPoint2dAt(c).Y)
              xmpoint.InnerText = t.ToString()
              xpline.AppendChild(xmpoint)
            Next
            If pline.Closed() Then
              Dim xmpoint0 As XmlElement = xmdoc.CreateElement("Point")
              Dim t0 As StringBuilder = New StringBuilder()
              t0.Append(pline.GetPoint2dAt(0).X)
              t0.Append(", ")
              t0.Append(pline.GetPoint2dAt(0).Y)
              xmpoint0.InnerText = t0.ToString()
              xpline.AppendChild(xmpoint0)
            End If
            layer.AppendChild(xpline)

          Case "Polyline2d"
            Dim pline As Polyline2d = tm.GetObject(Id, OpenMode.ForRead)
            Dim xpline As XmlElement = xmdoc.CreateElement("Polyline")
            For c As Integer = 0 To Convert.ToInt32(pline.EndParam())
              Dim xmpoint As XmlElement = xmdoc.CreateElement("Point")
              Dim t As StringBuilder = New StringBuilder()
              t.Append(pline.GetPointAtParameter(c).X)
              t.Append(", ")
              t.Append(pline.GetPointAtParameter(c).Y)
              xmpoint.InnerText = t.ToString()
              xpline.AppendChild(xmpoint)
            Next
            If pline.Closed() Then
              Dim xmpoint0 As XmlElement = xmdoc.CreateElement("Point")
              Dim t0 As StringBuilder = New StringBuilder()
              t0.Append(pline.GetPointAtParameter(0).X)
              t0.Append(", ")
              t0.Append(pline.GetPointAtParameter(0).Y)
              xmpoint0.InnerText = t0.ToString()
              xpline.AppendChild(xmpoint0)
            End If
            layer.AppendChild(xpline)


          Case "BlockReference"
            Dim hn As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
            If hn.Name = "hnumu" Then
              Dim xmHouseNumber As XmlElement = xmdoc.CreateElement("HouseNumber")
              xmHouseNumber.SetAttribute("handle", hn.Handle.ToString())
              Dim xmNumber As XmlElement = xmdoc.CreateElement("Number")
              Dim xmAddInfo As XmlElement = xmdoc.CreateElement("AddInfo")
              Dim xmStreet As XmlElement = xmdoc.CreateElement("StreetName")
              Dim xmOid As XmlElement = xmdoc.CreateElement("Oid")
              Dim xmZipCode As XmlElement = xmdoc.CreateElement("ZipCode")
              Dim xmPointX As XmlElement = xmdoc.CreateElement("X")
              Dim xmPointY As XmlElement = xmdoc.CreateElement("Y")
              Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
              xmPointX.InnerText = hn.Position.X.ToString()
              xmPointY.InnerText = hn.Position.Y.ToString()
              xmRotation.InnerText = hn.Rotation.ToString()
              Dim rs As Array = hn.GetXDataForApplication("ESTATE_DATAS").AsArray()
              xmStreet.InnerText = rs(1).Value
              xmOid.InnerText = rs(2).Value
              xmNumber.InnerText = rs(3).Value
              xmAddInfo.InnerText = rs(4).Value
              xmZipCode.InnerText = rs(11).Value
              xmHouseNumber.AppendChild(xmOid)
              xmHouseNumber.AppendChild(xmNumber)
              xmHouseNumber.AppendChild(xmAddInfo)
              xmHouseNumber.AppendChild(xmStreet)
              xmHouseNumber.AppendChild(xmZipCode)
              xmHouseNumber.AppendChild(xmPointX)
              xmHouseNumber.AppendChild(xmPointY)
              xmHouseNumber.AppendChild(xmRotation)
              layer.AppendChild(xmHouseNumber)
            End If
          Case "Circle"
            Dim Cir As Circle = tm.GetObject(Id, OpenMode.ForRead)
            Dim xmCircle As XmlElement = xmdoc.CreateElement("Circle")
            xmCircle.SetAttribute("handle", Cir.Handle.ToString())
            Dim xmPointX As XmlElement = xmdoc.CreateElement("X")
            Dim xmPointY As XmlElement = xmdoc.CreateElement("Y")
            Dim xmRadius As XmlElement = xmdoc.CreateElement("Radius")
            xmPointX.InnerText = Cir.Center.X.ToString()
            xmPointY.InnerText = Cir.Center.Y.ToString()
            xmRadius.InnerText = Cir.Radius.ToString()
            xmCircle.AppendChild(xmPointX)
            xmCircle.AppendChild(xmPointY)
            xmCircle.AppendChild(xmRadius)
            layer.AppendChild(xmCircle)
        End Select
      Next

      'Запись данных в XML
      'layer.SetAttribute("Layer", "Houses")
      'layer.AppendChild(layer)
    Finally
      ''
    End Try
    Return layer
  End Function


  Public Function ExportCable(ByVal ed As Editor) As XmlElement
    Dim xmCables As XmlElement = xmdoc.CreateElement("Cables")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXCABLE,PLCPOWERCABLE"), _
                New TypedValue(DxfCode.Start, "POLYLINE")}

    Dim sfilter As New SelectionFilter(values)

    Dim Res As PromptSelectionResult
    'Dim selopts As PromptSelectionOptions = New PromptSelectionOptions()
    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт кабелей.")
    ed.UpdateScreen()
    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о кабелях")
      ed.UpdateScreen()
      Return xmCables
    End If

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim xmCable As XmlElement = xmdoc.CreateElement("Cable")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmPartName As XmlElement = xmdoc.CreateElement("PartName")
      Dim xmLayer As XmlElement = xmdoc.CreateElement("Layer")
      Dim xmLength As XmlElement = xmdoc.CreateElement("Length")

      Dim pent As Entity = tm.GetObject(Id, OpenMode.ForRead)

      '      Dim s As String = Id.GetType().Name.ToString()

      Select Case pent.GetType().Name.ToString()
        Case "Polyline2d"
          Dim pline As Polyline2d = tm.GetObject(Id, OpenMode.ForRead)

          Dim rsParentCross As ResultBuffer = pline.GetXDataForApplication("PARENT_CROSS")
          If rsParentCross = Nothing Then
            xmCable.SetAttribute("PARENT_CROSS", "NONE")
          Else
            Dim arParentCross As Array = rsParentCross.AsArray()
            xmCable.SetAttribute("PARENT_CROSS", arParentCross(1).Value)
          End If

          xmHandle.InnerText = pline.Handle.ToString()
          Dim rs2 As ResultBuffer = pline.GetXDataForApplication("DEVICE_DATAS")
          If rs2 Is Nothing Then
            ed.WriteMessage("Проблема с экспортом кабелей!!!!!!!!!!!!!!!!!!!!")
          Else
            Dim rs As Array = rs2.AsArray()



            xmPartName.InnerText = rs(8).Value
            xmLayer.InnerText = rs(4).Value
            xmLength.InnerText = rs(38).Value

            xmCable.AppendChild(xmHandle)
            xmCable.AppendChild(xmPartName)
            xmCable.AppendChild(xmLength)
            xmCable.AppendChild(xmLayer)

          End If

          For c As Integer = 0 To Convert.ToInt32(pline.EndParam())
            Dim xmpoint As XmlElement = xmdoc.CreateElement("Point")
            Dim t As StringBuilder = New StringBuilder()
            t.Append(pline.GetPointAtParameter(c).X)
            t.Append(", ")
            t.Append(pline.GetPointAtParameter(c).Y)
            xmpoint.InnerText = t.ToString()
            xmCable.AppendChild(xmpoint)
          Next
          xmCables.AppendChild(xmCable)
        Case "Line"
          '     Dim Line As Line = tm.GetObject(Id, OpenMode.ForWrite, True)
          '      Line.Erase()
      End Select
    Next

    Return xmCables
  End Function

  Public Function ExportCableMarkers(ByVal ed As Editor) As XmlElement

    Dim xmCabTypes As XmlElement = xmdoc.CreateElement("CableMarkers")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCABLELABEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт кабельных меток.")
    ed.UpdateScreen()
    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о метках кабелей")
      ed.UpdateScreen()
      Return xmCabTypes
    End If

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmCabType As XmlElement = xmdoc.CreateElement("CableMarker")
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")

      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          xmCabType.AppendChild(xmBezRef)
        Next
      End If

      Dim xmType As XmlElement = xmdoc.CreateElement("CAB_TYP")
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        If attref.Tag = "TYP" Then
          xmType.InnerText = attref.TextString
          xmCabType.AppendChild(xmType)
        End If
      Next
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmCabType.AppendChild(xmX)
      xmCabType.AppendChild(xmY)
      xmCabType.AppendChild(xmRotation)
      xmCabTypes.AppendChild(xmCabType)
    Next

    Return xmCabTypes
  End Function

  Public Function ExportAmplifiers(ByVal ed As Editor) As XmlElement
    Dim xmAmplifiers As XmlElement = xmdoc.CreateElement("Amplifiers")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXAMPLIFIER")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    'Res = ed.GetSelection(SelOpts, sfilter)
    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт усилителей.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)

    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об усилителях")
      ed.UpdateScreen()
      Return xmAmplifiers
    End If

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length



    Dim Id As ObjectId
    For Each Id In IdArray

      Dim ks As Entity = tm.GetObject(Id, OpenMode.ForRead)

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      'Dim Ent As Entity = tm.GetObject(Id, OpenMode.ForRead)


      Dim xmAmplifier As XmlElement = xmdoc.CreateElement("Amplifier")
      Dim rsParentCross As ResultBuffer = blkRef.GetXDataForApplication("PARENT_CROSS")

      If rsParentCross = Nothing Then
        xmAmplifier.SetAttribute("PARENT_CROSS", "NONE")
      Else
        Dim arParentCross As Array = rsParentCross.AsArray()
        xmAmplifier.SetAttribute("PARENT_CROSS", arParentCross(1).Value)
      End If

      Dim xmParentAmp As XmlElement = xmdoc.CreateElement("ParentAmplifier")
      xmAmplifier.AppendChild(xmParentAmp)

      'Вывод данных об аттрибуте
      Dim xmlAttribute As XmlElement = xmdoc.CreateElement("BlockAttribute")
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        Dim xmlX As XmlElement = xmdoc.CreateElement("X")
        Dim xmlY As XmlElement = xmdoc.CreateElement("Y")
        Dim xmlRotation As XmlElement = xmdoc.CreateElement("Rotation")
        Dim xmlText As XmlElement = xmdoc.CreateElement("Text")
        xmlX.InnerText = attref.Position.X.ToString()
        xmlY.InnerText = attref.Position.Y.ToString()
        xmlRotation.InnerText = attref.Rotation.ToString()
        xmlText.InnerText = attref.TextString()
        xmlAttribute.AppendChild(xmlX)
        xmlAttribute.AppendChild(xmlY)
        xmlAttribute.AppendChild(xmlRotation)
        xmlAttribute.AppendChild(xmlText)
      Next
      xmAmplifier.AppendChild(xmlAttribute)

      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmBlockName As XmlElement = xmdoc.CreateElement("BlockName")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmBlockName.InnerText = blkRef.Name
      xmAmplifier.AppendChild(xmX)
      xmAmplifier.AppendChild(xmY)
      xmAmplifier.AppendChild(xmRotation)
      xmAmplifier.AppendChild(xmHandle)
      xmAmplifier.AppendChild(xmBlockName)

      Dim Width As XmlElement = xmdoc.CreateElement("Width")
      Dim Height As XmlElement = xmdoc.CreateElement("Height")
      Dim h, w As Double
      Select Case blkRef.Name
        Case "AMP1_1", "AMP1_19"
          h = 6
          w = 6
        Case "AMP1_6"
          h = 6
          w = 12
        Case "AMP1_18"
          h = 6
          w = 9
        Case "AMP1_20"
          h = 10
          w = 5
        Case Else
          h = 6
          w = 6
      End Select
      Width.InnerText = w * blkRef.ScaleFactors.X
      Height.InnerText = h * blkRef.ScaleFactors.Y
      xmAmplifier.AppendChild(Width)
      xmAmplifier.AppendChild(Height)

      'Ссылка на маркер усилителя
      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          ''xmCabType.AppendChild(xmBezRef)
          xmAmplifier.AppendChild(xmBezRef)
        Next
      Else
        Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
        xmBezRef.InnerText = "NOTHING"
        xmAmplifier.AppendChild(xmBezRef)
      End If
      'Вывод данных об усилителе
      Dim xmDeviceDatas As XmlElement = xmdoc.CreateElement("DeviceDatas")
      Dim rsDevData As Array = blkRef.GetXDataForApplication("DEVICE_DATAS").AsArray()
      If rsDevData(1).Value = "DEVICE_ID" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Device_ID")
        xm.InnerText = rsDevData(2).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(5).Value = "PART_NAME" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Part_name")
        xm.InnerText = rsDevData(6).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(7).Value = "DESCRIPTION" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Description")
        xm.InnerText = rsDevData(8).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      xmAmplifier.AppendChild(xmDeviceDatas)

      'Запись адреса усилителя
      Dim rsEstate As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()
      Dim xmEstate As XmlElement = xmdoc.CreateElement("EstateDatas")
      Dim xmNumber As XmlElement = xmdoc.CreateElement("Number")
      Dim xmAddInfo As XmlElement = xmdoc.CreateElement("AddInfo")
      Dim xmStreet As XmlElement = xmdoc.CreateElement("StreetName")
      Dim xmOid As XmlElement = xmdoc.CreateElement("Oid")
      Dim xmZipCode As XmlElement = xmdoc.CreateElement("ZipCode")

      Dim rs1 As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()

      If Not rs1(1).Value.ToString().Length = 0 Then
        xmStreet.InnerText = rs1(1).Value
        xmOid.InnerText = FindFid(rs1(3).Value, rs1(4).Value, rs1(1).Value, rs1(11).value)
        xmNumber.InnerText = rs1(3).Value
        xmAddInfo.InnerText = rs1(4).Value
        xmZipCode.InnerText = rs1(11).Value
        xmEstate.AppendChild(xmOid)
        xmEstate.AppendChild(xmNumber)
        xmEstate.AppendChild(xmAddInfo)
        xmEstate.AppendChild(xmStreet)
        xmEstate.AppendChild(xmZipCode)
        xmAmplifier.AppendChild(xmEstate)
      Else
        Dim s As StringBuilder = New StringBuilder()
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage("Усилитель не имеет адреса:")
        ed.WriteMessage(ControlChars.CrLf)
        s.Append("Координаты x = ")
        s.Append(blkRef.Position.X)
        s.Append(" ;y = ")
        s.Append(blkRef.Position.Y)
        ed.WriteMessage(s.ToString())
        ed.WriteMessage(ControlChars.CrLf)
        ed.UpdateScreen()
      End If

      xmAmplifiers.AppendChild(xmAmplifier)
    Next

    Return xmAmplifiers
  End Function



  Public Function ExportTaps(ByVal ed As Editor) As XmlElement
    Dim xmTaps As XmlElement = xmdoc.CreateElement("Taps")

    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXTAP")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт ответвителей.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об ответвителях.")
      ed.UpdateScreen()
      Return xmTaps
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)

      Dim xmTap As XmlElement = xmdoc.CreateElement("Tap")
      Dim rsParentCross As ResultBuffer = blkRef.GetXDataForApplication("PARENT_CROSS")

      If rsParentCross = Nothing Then
        xmTap.SetAttribute("PARENT_CROSS", "NONE")
      Else
        Dim arParentCross As Array = rsParentCross.AsArray()
        xmTap.SetAttribute("PARENT_CROSS", arParentCross(1).Value)
      End If


      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmBlockName As XmlElement = xmdoc.CreateElement("BlockName")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmBlockName.InnerText = blkRef.Name
      xmTap.AppendChild(xmX)
      xmTap.AppendChild(xmY)
      xmTap.AppendChild(xmRotation)
      xmTap.AppendChild(xmHandle)
      xmTap.AppendChild(xmBlockName)

      Dim Width As XmlElement = xmdoc.CreateElement("Width")
      Dim Height As XmlElement = xmdoc.CreateElement("Height")
      Dim h, w As Double
      Select Case blkRef.Name
        Case "TAP1_1", "TAP1_2", "TAP1_17", "TAP1_18", "TAP1_21", "TAP1_22", _
          "TAP2_2", "TAP2_3", "TAP3_3"
          h = 5
          w = 6
        Case "TAP2_1", "TAP2_13", "TAP3_1", "TAP3_2", "TAP4_1", "TAP4_9", _
        "TAP4_12", "TAP6_1", "TAP6_4"
          h = 6
          w = 6
        Case "TAP1_7", "TAP1_8"
          h = 4
          w = 2.5
        Case "TAP1_9"
          h = 4
          w = 4
        Case "TAP1_19"
          h = 4
          w = 3
        Case "TAP1_20"
          h = 3
          w = 2
        Case Else
          h = 5
          w = 6
      End Select
      Width.InnerText = w * blkRef.ScaleFactors.X
      Height.InnerText = h * blkRef.ScaleFactors.Y
      xmTap.AppendChild(Width)
      xmTap.AppendChild(Height)

      'Вывод данных об аттрибуте
      Dim xmlAttribute As XmlElement = xmdoc.CreateElement("BlockAttribute")
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        Dim xmlX As XmlElement = xmdoc.CreateElement("X")
        Dim xmlY As XmlElement = xmdoc.CreateElement("Y")
        Dim xmlRotation As XmlElement = xmdoc.CreateElement("Rotation")
        Dim xmlText As XmlElement = xmdoc.CreateElement("Text")
        xmlX.InnerText = attref.Position.X.ToString()
        xmlY.InnerText = attref.Position.Y.ToString()
        xmlRotation.InnerText = attref.Rotation.ToString()
        xmlText.InnerText = attref.TextString()
        xmlAttribute.AppendChild(xmlX)
        xmlAttribute.AppendChild(xmlY)
        xmlAttribute.AppendChild(xmlRotation)
        xmlAttribute.AppendChild(xmlText)
      Next
      xmTap.AppendChild(xmlAttribute)

      'Ссылка на маркер ответвителя
      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          ''xmCabType.AppendChild(xmBezRef)
          xmTap.AppendChild(xmBezRef)
        Next
      Else
        Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
        xmBezRef.InnerText = "NOTHING"
        xmTap.AppendChild(xmBezRef)
      End If
      'Вывод данных об ответвителе
      Dim xmDeviceDatas As XmlElement = xmdoc.CreateElement("DeviceDatas")
      Dim rsDevData As Array = blkRef.GetXDataForApplication("DEVICE_DATAS").AsArray()
      If rsDevData(1).Value = "DEVICE_ID" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Device_ID")
        xm.InnerText = rsDevData(2).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(7).Value = "PART_NAME" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Part_name")
        xm.InnerText = rsDevData(8).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(17).Value = "DESCRIPTION" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Description")
        xm.InnerText = rsDevData(18).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      xmTap.AppendChild(xmDeviceDatas)

      'Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      'Dim attid As ObjectId
      'For Each attid In attcol
      'Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
      'If attref.Tag = "TYP" Then
      'xmType.InnerText = attref.TextString
      'xmCabType.AppendChild(xmType)
      'End If
      'Next
      'Запись адреса усилителя
      Dim rsEstate As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()
      Dim xmEstate As XmlElement = xmdoc.CreateElement("EstateDatas")
      Dim xmNumber As XmlElement = xmdoc.CreateElement("Number")
      Dim xmAddInfo As XmlElement = xmdoc.CreateElement("AddInfo")
      Dim xmStreet As XmlElement = xmdoc.CreateElement("StreetName")
      Dim xmOid As XmlElement = xmdoc.CreateElement("Oid")
      Dim xmZipCode As XmlElement = xmdoc.CreateElement("ZipCode")

      Dim rs1 As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()
      If Not rs1(1).Value.ToString().Length = 0 Then
        xmStreet.InnerText = rs1(1).Value
        xmOid.InnerText = FindFid(rs1(3).Value, rs1(4).Value, rs1(1).Value, rs1(11).value)
        xmNumber.InnerText = rs1(3).Value
        xmAddInfo.InnerText = rs1(4).Value
        xmZipCode.InnerText = rs1(11).Value
        xmEstate.AppendChild(xmOid)
        xmEstate.AppendChild(xmNumber)
        xmEstate.AppendChild(xmAddInfo)
        xmEstate.AppendChild(xmStreet)
        xmEstate.AppendChild(xmZipCode)
        xmTap.AppendChild(xmEstate)
      Else
        Dim s As StringBuilder = New StringBuilder()
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage("Ответвитель не имеет адреса:")
        ed.WriteMessage(ControlChars.CrLf)
        s.Append("Координаты x = ")
        s.Append(blkRef.Position.X)
        s.Append(" ;y = ")
        s.Append(blkRef.Position.Y)
        ed.WriteMessage(s.ToString())
        ed.WriteMessage(ControlChars.CrLf)
        ed.UpdateScreen()
      End If

      xmTaps.AppendChild(xmTap)
    Next
    Return xmTaps
  End Function

  Public Function ExportSplitters(ByVal ed As Editor) As XmlElement
    Dim xmSplitters As XmlElement = xmdoc.CreateElement("Splitters")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXSPLITTER")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт делителей.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о делителях.")
      ed.UpdateScreen()
      Return xmSplitters
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)

      Dim xmSplitter As XmlElement = xmdoc.CreateElement("Splitter")
      Dim rsParentCross As ResultBuffer = blkRef.GetXDataForApplication("PARENT_CROSS")

      If rsParentCross = Nothing Then
        xmSplitter.SetAttribute("PARENT_CROSS", "NONE")
      Else
        Dim arParentCross As Array = rsParentCross.AsArray()
        xmSplitter.SetAttribute("PARENT_CROSS", arParentCross(1).Value)
      End If

      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmBlockName As XmlElement = xmdoc.CreateElement("BlockName")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmBlockName.InnerText = blkRef.Name
      xmSplitter.AppendChild(xmX)
      xmSplitter.AppendChild(xmY)
      xmSplitter.AppendChild(xmRotation)
      xmSplitter.AppendChild(xmHandle)
      xmSplitter.AppendChild(xmBlockName)

      Dim Width As XmlElement = xmdoc.CreateElement("Width")
      Dim Height As XmlElement = xmdoc.CreateElement("Height")
      Dim h, w As Double
      Select Case blkRef.Name
        Case "SPL2_1", "SPL2_2", "SPL2_3"
          h = 5
          w = 6
        Case "SPL3_1", "SPL4_1", "SPL6_1"
          h = 6
          w = 6
        Case "SPL2_9"
          h = 4
          w = 4
        Case Else
          h = 5
          w = 6
      End Select
      Width.InnerText = w * blkRef.ScaleFactors.X
      Height.InnerText = h * blkRef.ScaleFactors.Y
      xmSplitter.AppendChild(Width)
      xmSplitter.AppendChild(Height)

      'Вывод данных об аттрибуте
      Dim xmlAttribute As XmlElement = xmdoc.CreateElement("BlockAttribute")
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        Dim xmlX As XmlElement = xmdoc.CreateElement("X")
        Dim xmlY As XmlElement = xmdoc.CreateElement("Y")
        Dim xmlRotation As XmlElement = xmdoc.CreateElement("Rotation")
        Dim xmlText As XmlElement = xmdoc.CreateElement("Text")
        xmlX.InnerText = attref.Position.X.ToString()
        xmlY.InnerText = attref.Position.Y.ToString()
        xmlRotation.InnerText = attref.Rotation.ToString()
        xmlText.InnerText = attref.TextString()
        xmlAttribute.AppendChild(xmlX)
        xmlAttribute.AppendChild(xmlY)
        xmlAttribute.AppendChild(xmlRotation)
        xmlAttribute.AppendChild(xmlText)
      Next
      xmSplitter.AppendChild(xmlAttribute)

      'Ссылка на маркер ответвителя
      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          ''xmCabType.AppendChild(xmBezRef)
          xmSplitter.AppendChild(xmBezRef)
        Next
      Else
        Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
        xmBezRef.InnerText = "NOTHING"
        xmSplitter.AppendChild(xmBezRef)
      End If
      'Вывод данных об усилителе
      Dim xmDeviceDatas As XmlElement = xmdoc.CreateElement("DeviceDatas")
      Dim rsDevData As Array = blkRef.GetXDataForApplication("DEVICE_DATAS").AsArray()
      If rsDevData(1).Value = "DEVICE_ID" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Device_ID")
        xm.InnerText = rsDevData(2).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(7).Value = "PART_NAME" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Part_name")
        xm.InnerText = rsDevData(8).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(15).Value = "DESCRIPTION" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Description")
        xm.InnerText = rsDevData(16).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      xmSplitter.AppendChild(xmDeviceDatas)

      'Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      'Dim attid As ObjectId
      'For Each attid In attcol
      'Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
      'If attref.Tag = "TYP" Then
      'xmType.InnerText = attref.TextString
      'xmCabType.AppendChild(xmType)
      'End If
      'Next
      'Запись адреса усилителя

      Dim rsEstate As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()
      Dim xmEstate As XmlElement = xmdoc.CreateElement("EstateDatas")
      Dim xmNumber As XmlElement = xmdoc.CreateElement("Number")
      Dim xmAddInfo As XmlElement = xmdoc.CreateElement("AddInfo")
      Dim xmStreet As XmlElement = xmdoc.CreateElement("StreetName")
      Dim xmOid As XmlElement = xmdoc.CreateElement("Oid")
      Dim xmZipCode As XmlElement = xmdoc.CreateElement("ZipCode")

      Dim rs1 As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()
      If Not rs1(1).Value.ToString().Length = 0 Then
        xmStreet.InnerText = rs1(1).Value
        xmOid.InnerText = FindFid(rs1(3).Value, rs1(4).Value, rs1(1).Value, rs1(11).value)
        xmNumber.InnerText = rs1(3).Value
        xmAddInfo.InnerText = rs1(4).Value
        xmZipCode.InnerText = rs1(11).Value
        xmEstate.AppendChild(xmOid)
        xmEstate.AppendChild(xmNumber)
        xmEstate.AppendChild(xmAddInfo)
        xmEstate.AppendChild(xmStreet)
        xmEstate.AppendChild(xmZipCode)
        xmSplitter.AppendChild(xmEstate)
      Else
        Dim s As StringBuilder = New StringBuilder()
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage("Делитель не имеет адреса:")
        ed.WriteMessage(ControlChars.CrLf)
        s.Append("Координаты x = ")
        s.Append(blkRef.Position.X)
        s.Append(" ;y = ")
        s.Append(blkRef.Position.Y)
        ed.WriteMessage(s.ToString())
        ed.WriteMessage(ControlChars.CrLf)
        ed.UpdateScreen()
      End If
      xmSplitters.AppendChild(xmSplitter)
    Next
    Return xmSplitters
  End Function


  Public Function ExportSignalPoint(ByVal ed As Editor) As XmlElement
    Dim xmSignalPoints As XmlElement = xmdoc.CreateElement("SignalPoints")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCSIGNALPOINT")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт источников сигналов.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об источниках сигналов")
      ed.UpdateScreen()
      Return xmSignalPoints
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmSignalPoint As XmlElement = xmdoc.CreateElement("SignalPoint")

      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmOpticalNode As XmlElement = xmdoc.CreateElement("OpticalNode")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmOpticalNode.InnerText = FindOpticalNode(ed, blkRef.Position)
      xmSignalPoint.AppendChild(xmX)
      xmSignalPoint.AppendChild(xmY)
      xmSignalPoint.AppendChild(xmRotation)
      xmSignalPoint.AppendChild(xmHandle)
      xmSignalPoint.AppendChild(xmOpticalNode)

      'Вывод данных об аттрибуте
      Dim xmlAttribute As XmlElement = xmdoc.CreateElement("BlockAttribute")
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        Dim xmlX As XmlElement = xmdoc.CreateElement("X")
        Dim xmlY As XmlElement = xmdoc.CreateElement("Y")
        Dim xmlRotation As XmlElement = xmdoc.CreateElement("Rotation")
        Dim xmlText As XmlElement = xmdoc.CreateElement("Text")
        xmlX.InnerText = attref.Position.X.ToString()
        xmlY.InnerText = attref.Position.Y.ToString()
        xmlRotation.InnerText = attref.Rotation.ToString()
        xmlText.InnerText = attref.TextString()
        xmlAttribute.AppendChild(xmlX)
        xmlAttribute.AppendChild(xmlY)
        xmlAttribute.AppendChild(xmlRotation)
        xmlAttribute.AppendChild(xmlText)
      Next
      xmSignalPoint.AppendChild(xmlAttribute)

      'Ссылка на маркер ответвителя
      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          xmSignalPoint.AppendChild(xmBezRef)
        Next
      Else
        Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
        xmBezRef.InnerText = "NOTHING"
        xmSignalPoint.AppendChild(xmBezRef)
      End If

      'Вывод данных об источнике сигнала
      Dim xmDeviceDatas As XmlElement = xmdoc.CreateElement("DeviceDatas")
      Dim rsDevData As Array = blkRef.GetXDataForApplication("DEVICE_DATAS").AsArray()
      If rsDevData(1).Value = "PROJECT_ID" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Project_ID")
        xm.InnerText = rsDevData(2).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(3).Value = "DRAWING_ID" Then
        Dim xm As XmlElement = xmdoc.CreateElement("Drawing_ID")
        xm.InnerText = rsDevData(4).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      If rsDevData(5).Value = "POINT_ID" Then
        Dim xm As XmlElement = xmdoc.CreateElement("PointID")
        xm.InnerText = rsDevData(6).Value
        xmDeviceDatas.AppendChild(xm)
      End If
      xmSignalPoint.AppendChild(xmDeviceDatas)

      'Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      'Dim attid As ObjectId
      'For Each attid In attcol
      'Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
      'If attref.Tag = "TYP" Then
      'xmType.InnerText = attref.TextString
      'xmCabType.AppendChild(xmType)
      'End If
      'Next

      xmSignalPoints.AppendChild(xmSignalPoint)
    Next
    Return xmSignalPoints
  End Function


  Public Function ExportOutlets(ByVal ed As Editor) As XmlElement
    Dim xmOutlets As XmlElement = xmdoc.CreateElement("Outlets")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXOUTLET")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт розеток.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о розетках")
      ed.UpdateScreen()
      Return xmOutlets
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmOutlet As XmlElement = xmdoc.CreateElement("Outlet")

      Dim rsParentCross As ResultBuffer = blkRef.GetXDataForApplication("PARENT_CROSS")
      If rsParentCross = Nothing Then
        xmOutlet.SetAttribute("PARENT_CROSS", "NONE")
      Else
        Dim arParentCross As Array = rsParentCross.AsArray()
        xmOutlet.SetAttribute("PARENT_CROSS", arParentCross(1).Value)
      End If

      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmBlockName As XmlElement = xmdoc.CreateElement("BlockName")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmBlockName.InnerText = blkRef.Name
      xmOutlet.AppendChild(xmX)
      xmOutlet.AppendChild(xmY)
      xmOutlet.AppendChild(xmRotation)
      xmOutlet.AppendChild(xmHandle)
      xmOutlet.AppendChild(xmBlockName)

      Dim Width As XmlElement = xmdoc.CreateElement("Width")
      Dim Height As XmlElement = xmdoc.CreateElement("Height")
      Dim h, w As Double
      Select Case blkRef.Name
        Case "PLG1_4-SW"
          h = 1.23
          w = 1.03
        Case Else
          h = 3
          w = 2
      End Select
      Width.InnerText = w * blkRef.ScaleFactors.X
      Height.InnerText = h * blkRef.ScaleFactors.Y
      xmOutlet.AppendChild(Width)
      xmOutlet.AppendChild(Height)

      'Вывод данных об аттрибуте
      Dim xmlAttribute As XmlElement = xmdoc.CreateElement("BlockAttribute")
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        Dim xmlX As XmlElement = xmdoc.CreateElement("X")
        Dim xmlY As XmlElement = xmdoc.CreateElement("Y")
        Dim xmlRotation As XmlElement = xmdoc.CreateElement("Rotation")
        Dim xmlText As XmlElement = xmdoc.CreateElement("Text")
        xmlX.InnerText = attref.Position.X.ToString()
        xmlY.InnerText = attref.Position.Y.ToString()
        xmlRotation.InnerText = attref.Rotation.ToString()
        xmlText.InnerText = attref.TextString()
        xmlAttribute.AppendChild(xmlX)
        xmlAttribute.AppendChild(xmlY)
        xmlAttribute.AppendChild(xmlRotation)
        xmlAttribute.AppendChild(xmlText)
      Next
      xmOutlet.AppendChild(xmlAttribute)

      'Ссылка на маркер ответвителя
      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          ''xmCabType.AppendChild(xmBezRef)
          xmOutlet.AppendChild(xmBezRef)
        Next
      Else
        Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
        xmBezRef.InnerText = "NOTHING"
        xmOutlet.AppendChild(xmBezRef)
      End If

      'Вывод данных об розетке
      Dim xmDeviceDatas As XmlElement = xmdoc.CreateElement("DeviceDatas")

      Dim rsDeviceDatas As ResultBuffer = blkRef.GetXDataForApplication("DEVICE_DATAS")

      If (rsDeviceDatas <> Nothing) Then
        Dim rsDevData As Array = rsDeviceDatas.AsArray()
        If rsDevData(1).Value = "DEVICE_ID" Then
          Dim xm As XmlElement = xmdoc.CreateElement("Device_ID")
          xm.InnerText = rsDevData(2).Value
          xmDeviceDatas.AppendChild(xm)
        End If
        If rsDevData(5).Value = "PART_NAME" Then
          Dim xm As XmlElement = xmdoc.CreateElement("Part_name")
          xm.InnerText = rsDevData(6).Value
          xmDeviceDatas.AppendChild(xm)
        End If
        If rsDevData(13).Value = "DESCRIPTION" Then
          Dim xm As XmlElement = xmdoc.CreateElement("Description")
          xm.InnerText = rsDevData(14).Value
          xmDeviceDatas.AppendChild(xm)
        End If
        xmOutlet.AppendChild(xmDeviceDatas)
      Else
        ed.WriteMessage("Ошибка в расширенных данных у розетки")
      End If


      'Запись адреса розетки
      Dim rsEstate As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()
      Dim xmEstate As XmlElement = xmdoc.CreateElement("EstateDatas")
      Dim xmNumber As XmlElement = xmdoc.CreateElement("Number")
      Dim xmAddInfo As XmlElement = xmdoc.CreateElement("AddInfo")
      Dim xmStreet As XmlElement = xmdoc.CreateElement("StreetName")
      Dim xmOid As XmlElement = xmdoc.CreateElement("Oid")
      Dim xmZipCode As XmlElement = xmdoc.CreateElement("ZipCode")

      Dim rs1 As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()
      xmStreet.InnerText = rs1(1).Value
      xmOid.InnerText = FindFid(rs1(3).Value, rs1(4).Value, rs1(1).Value, rs1(11).value)
      xmNumber.InnerText = rs1(3).Value
      xmAddInfo.InnerText = rs1(4).Value
      xmZipCode.InnerText = rs1(11).Value
      xmEstate.AppendChild(xmOid)
      xmEstate.AppendChild(xmNumber)
      xmEstate.AppendChild(xmAddInfo)
      xmEstate.AppendChild(xmStreet)
      xmEstate.AppendChild(xmZipCode)
      xmOutlet.AppendChild(xmEstate)

      xmOutlets.AppendChild(xmOutlet)
    Next
    Return xmOutlets
  End Function

  Public Function ExportPowerSources(ByVal ed As Editor) As XmlElement
    Dim xmPowerSources As XmlElement = xmdoc.CreateElement("PowerSources")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCREMOTEPOWERSOURCE")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    'Res = ed.GetSelection(SelOpts, sfilter)
    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт блоков питания.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)

    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о блоках питания.")
      ed.UpdateScreen()
      Return xmPowerSources
    End If

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      'Dim Ent As Entity = tm.GetObject(Id, OpenMode.ForRead)


      Dim xmPowerSource As XmlElement = xmdoc.CreateElement("PowerSource")

      'Вывод данных об аттрибуте
      Dim xmlAttribute As XmlElement = xmdoc.CreateElement("BlockAttribute")
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        Dim xmlX As XmlElement = xmdoc.CreateElement("X")
        Dim xmlY As XmlElement = xmdoc.CreateElement("Y")
        Dim xmlRotation As XmlElement = xmdoc.CreateElement("Rotation")
        Dim xmlText As XmlElement = xmdoc.CreateElement("Text")
        xmlX.InnerText = attref.Position.X.ToString()
        xmlY.InnerText = attref.Position.Y.ToString()
        xmlRotation.InnerText = attref.Rotation.ToString()
        xmlText.InnerText = attref.TextString()
        xmlAttribute.AppendChild(xmlX)
        xmlAttribute.AppendChild(xmlY)
        xmlAttribute.AppendChild(xmlRotation)
        xmlAttribute.AppendChild(xmlText)
      Next
      xmPowerSource.AppendChild(xmlAttribute)

      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmBlockName As XmlElement = xmdoc.CreateElement("BlockName")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmBlockName.InnerText = blkRef.Name
      xmPowerSource.AppendChild(xmX)
      xmPowerSource.AppendChild(xmY)
      xmPowerSource.AppendChild(xmRotation)
      xmPowerSource.AppendChild(xmHandle)
      xmPowerSource.AppendChild(xmBlockName)

      Dim Width As XmlElement = xmdoc.CreateElement("Width")
      Dim Height As XmlElement = xmdoc.CreateElement("Height")

      Width.InnerText = 11 * blkRef.ScaleFactors.X
      Height.InnerText = 1.6 * blkRef.ScaleFactors.Y
      xmPowerSource.AppendChild(Width)
      xmPowerSource.AppendChild(Height)


      'Вывод данных об усилителе
      '      Dim xmDeviceDatas As XmlElement = xmdoc.CreateElement("DeviceDatas")
      '     Dim rsDevData As Array = blkRef.GetXDataForApplication("DEVICE_DATAS").AsArray()
      '    If rsDevData(1).Value = "DEVICE_ID" Then
      'Dim xm As XmlElement = xmdoc.CreateElement("Device_ID")
      'xm.InnerText = rsDevData(2).Value
      'xmDeviceDatas.AppendChild(xm)
      'End If
      'If rsDevData(5).Value = "PART_NAME" Then
      ' Dim xm As XmlElement = xmdoc.CreateElement("Part_name")
      'xm.InnerText = rsDevData(6).Value
      'xmDeviceDatas.AppendChild(xm)
      'End If
      'If rsDevData(7).Value = "DESCRIPTION" Then
      'Dim xm As XmlElement = xmdoc.CreateElement("Description")
      'xm.InnerText = rsDevData(7).Value
      'xmDeviceDatas.AppendChild(xm)
      'End If
      'xmAmplifier.AppendChild(xmDeviceDatas)

      'Запись адреса блока питания

      Dim rsEstate As ResultBuffer = blkRef.GetXDataForApplication("ESTATE_DATAS")
      Dim xmEstate As XmlElement = xmdoc.CreateElement("EstateDatas")
      Dim xmNumber As XmlElement = xmdoc.CreateElement("Number")
      Dim xmAddInfo As XmlElement = xmdoc.CreateElement("AddInfo")
      Dim xmStreet As XmlElement = xmdoc.CreateElement("StreetName")
      Dim xmOid As XmlElement = xmdoc.CreateElement("Oid")
      Dim xmZipCode As XmlElement = xmdoc.CreateElement("ZipCode")
      If rsEstate <> Nothing Then
        Dim ArEstate As Array = rsEstate.AsArray()
        If Not ArEstate(1).Value.ToString().Length = 0 Then
          xmStreet.InnerText = ArEstate(1).Value
          xmOid.InnerText = FindFid(ArEstate(3).Value, ArEstate(4).Value, ArEstate(1).Value, ArEstate(11).value)
          xmNumber.InnerText = ArEstate(3).Value
          xmAddInfo.InnerText = ArEstate(4).Value
          xmZipCode.InnerText = ArEstate(11).Value
          xmEstate.AppendChild(xmOid)
          xmEstate.AppendChild(xmNumber)
          xmEstate.AppendChild(xmAddInfo)
          xmEstate.AppendChild(xmStreet)
          xmEstate.AppendChild(xmZipCode)
          xmPowerSource.AppendChild(xmEstate)
        Else
          Dim s As StringBuilder = New StringBuilder()
          ed.WriteMessage(ControlChars.CrLf)
          ed.WriteMessage(ControlChars.CrLf)
          ed.WriteMessage("Блок питания не имеет адресной информации:")
          ed.WriteMessage(ControlChars.CrLf)
          s.Append("Координаты x = ")
          s.Append(blkRef.Position.X)
          s.Append(" ;y = ")
          s.Append(blkRef.Position.Y)
          ed.WriteMessage(s.ToString())
          ed.WriteMessage(ControlChars.CrLf)
          ed.UpdateScreen()
        End If
        xmPowerSources.AppendChild(xmPowerSource)
      Else
        Dim s As StringBuilder = New StringBuilder()
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage(ControlChars.CrLf)
        ed.WriteMessage("Блок питания не имеет адреса:")
        ed.WriteMessage(ControlChars.CrLf)
        s.Append("Координаты x = ")
        s.Append(blkRef.Position.X)
        s.Append(" ;y = ")
        s.Append(blkRef.Position.Y)
        ed.WriteMessage(s.ToString())
        ed.WriteMessage(ControlChars.CrLf)
        ed.UpdateScreen()
      End If
    Next

    Return xmPowerSources
  End Function

  Public Function ExportSignalConnections(ByVal ed As Editor) As XmlElement
    Dim xmSignalConnections As XmlElement = xmdoc.CreateElement("SignalConnections")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "SIGNAL_CONNECTION")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт соединителей сигналов.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о входах/выходах сигналов")
      ed.UpdateScreen()
      Return xmSignalConnections
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmSignalConnection As XmlElement = xmdoc.CreateElement("SignalConnection")


      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmBlockName As XmlElement = xmdoc.CreateElement("BlockName")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmBlockName.InnerText = blkRef.Name
      xmSignalConnection.AppendChild(xmX)
      xmSignalConnection.AppendChild(xmY)
      xmSignalConnection.AppendChild(xmRotation)
      xmSignalConnection.AppendChild(xmHandle)
      xmSignalConnection.AppendChild(xmBlockName)

      Dim Width As XmlElement = xmdoc.CreateElement("Width")
      Dim Height As XmlElement = xmdoc.CreateElement("Height")
      Width.InnerText = 4.5 * blkRef.ScaleFactors.X
      Height.InnerText = 4.5 * blkRef.ScaleFactors.Y
      xmSignalConnection.AppendChild(Width)
      xmSignalConnection.AppendChild(Height)


      'Ссылка на маркер входа/выхода сигнала
      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          ''xmCabType.AppendChild(xmBezRef)
          xmSignalConnection.AppendChild(xmBezRef)
        Next
      Else
        Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
        xmBezRef.InnerText = "NOTHING"
        xmSignalConnection.AppendChild(xmBezRef)
      End If

      'Сссылка на соединитель сигнала
      Dim xmRef As XmlElement = xmdoc.CreateElement("SignalConnection")
      Dim rsDevData As Array = blkRef.GetXDataForApplication("SIGNAL_CONNECTION").AsArray()
      Dim xm As XmlElement = xmdoc.CreateElement("Handle")
      xm.InnerText = rsDevData(1).Value
      xmRef.AppendChild(xm)
      xmSignalConnection.AppendChild(xmRef)

      'Запись названия соединителя сигнала
      Dim rsSC As ResultBuffer = blkRef.GetXDataForApplication("PE_URL")
      If rsSC <> Nothing Then
        Dim arNameSC As Array = rsSC.AsArray()
        Dim xmNameSC As XmlElement = xmdoc.CreateElement("UrlSignalConnection")
        xmNameSC.InnerText = arNameSC(3).Value
        xmSignalConnection.AppendChild(xmNameSC)
      Else
        Dim xmNameSC As XmlElement = xmdoc.CreateElement("UrlSignalConnection")
        xmNameSC.InnerText = "None"
        xmSignalConnection.AppendChild(xmNameSC)
      End If
      xmSignalConnections.AppendChild(xmSignalConnection)
    Next
    Return xmSignalConnections
  End Function

  Public Function ExportTerminators(ByVal ed As Editor) As XmlElement
    Dim xmTerminators As XmlElement = xmdoc.CreateElement("Terminators")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXTERMINATOR")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    'Res = ed.GetSelection(SelOpts, sfilter)
    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт нагрузок 75 Ом.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)

    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об нагрузках ...")
      ed.UpdateScreen()
      Return xmTerminators
    End If

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray

      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      'Dim Ent As Entity = tm.GetObject(Id, OpenMode.ForRead)


      Dim xmTerminator As XmlElement = xmdoc.CreateElement("Terminator")

      Dim rsParentCross As ResultBuffer = blkRef.GetXDataForApplication("PARENT_CROSS")
      If rsParentCross = Nothing Then
        xmTerminator.SetAttribute("PARENT_CROSS", "NONE")
      Else
        Dim arParentCross As Array = rsParentCross.AsArray()
        xmTerminator.SetAttribute("PARENT_CROSS", arParentCross(1).Value)
      End If

      Dim xmParentAmp As XmlElement = xmdoc.CreateElement("ParentAmplifier")
      xmTerminator.AppendChild(xmParentAmp)

      'Запись геометрических параметров
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmHandle As XmlElement = xmdoc.CreateElement("Handle")
      Dim xmBlockName As XmlElement = xmdoc.CreateElement("BlockName")
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmHandle.InnerText = blkRef.Handle.ToString()
      xmBlockName.InnerText = blkRef.Name
      xmTerminator.AppendChild(xmX)
      xmTerminator.AppendChild(xmY)
      xmTerminator.AppendChild(xmRotation)
      xmTerminator.AppendChild(xmHandle)
      xmTerminator.AppendChild(xmBlockName)

      Dim Width As XmlElement = xmdoc.CreateElement("Width")
      Dim Height As XmlElement = xmdoc.CreateElement("Height")
      Width.InnerText = 2 * blkRef.ScaleFactors.X
      Height.InnerText = blkRef.ScaleFactors.Y
      xmTerminator.AppendChild(Width)
      xmTerminator.AppendChild(Height)

      xmTerminators.AppendChild(xmTerminator)
    Next
    Return xmTerminators
  End Function

  Public Function Analyse(ByVal doc As XmlElement) As XmlElement
    Dim Amplifiers As XmlNodeList
    Dim Amplifier As XmlElement
    Amplifier = doc.Item("Amplifiers")
    Amplifiers = doc.SelectNodes("descendant::Amplifier")
    Dim Amp As XmlElement
    For Each Amp In Amplifiers
      Dim s As String
      s = Amp.GetAttribute("PARENT_CROSS")
      If s = "NONE" Then
        Dim xm As XmlNodeList = Amp.ChildNodes
        Dim temp As XmlElement
        For Each temp In xm
          If temp.Name = "ParentAmplifier" Then
            temp.InnerText = "None"
            Dim oldAmp As XmlElement = Amp
            Amp.AppendChild(temp)
            Amplifier.ReplaceChild(Amp, oldAmp)
            'oldAmp.RemoveAll()
          End If
        Next
      Else
        Dim xm As XmlNodeList = Amp.ChildNodes
        Dim temp As XmlElement
        For Each temp In xm
          If temp.Name = "ParentAmplifier" Then
            temp.InnerText = FindParentAmplifier(s, doc)
            '            Dim sss As String
            '            sss = FindParentAmplifier2(s, doc)

            Dim oldAmp As XmlElement = Amp
            Amp.AppendChild(temp)
            Amplifier.ReplaceChild(Amp, oldAmp)
            'oldAmp.RemoveAll()
          End If
        Next
      End If
    Next
    Return doc
  End Function

  Public Function FindFid(ByVal Number As String, ByVal AddInfo As String, ByVal Street As String, ByVal zipcode As String) As Long
    Dim ad As EstateDatas

    For Each ad In Address
      If Number = ad.Number Then
        If AddInfo = ad.AddInfo Then
          If Street = ad.Street Then
            If zipcode = ad.ZipCode Then
              Return ad.Fid
            End If
          End If
        End If
      End If
    Next
    Return -1
  End Function

  Public Function FindParentAmplifier(ByVal Handle As String, ByVal doc As XmlElement) As String
    Dim xpath As StringBuilder = New StringBuilder()
    Dim xmlAmps As XmlElement
    Dim xmlTaps As XmlElement
    Dim xmlSplitters As XmlElement
    Dim xmlCabels As XmlElement
    Dim xmlSigPoints As XmlElement
    Dim s As String

    xpath.Append("descendant::Cable[Handle='")
    xpath.Append(Handle)
    xpath.Append("']")
    Dim ss As String = xpath.ToString()
    xmlCabels = doc.Item("Cables").SelectSingleNode(xpath.ToString())
    '  elements.RemoveAll()
    If xmlCabels Is Nothing Then
      xpath = New StringBuilder()
      xpath.Append("descendant::Tap[Handle='")
      xpath.Append(Handle)
      xpath.Append("']")
      xmlTaps = doc.Item("Taps").SelectSingleNode(xpath.ToString())
      If xmlTaps Is Nothing Then
        xpath = New StringBuilder()
        xpath.Append("descendant::Splitter[Handle='")
        xpath.Append(Handle)
        xpath.Append("']")
        xmlSplitters = doc.Item("Splitters").SelectSingleNode(xpath.ToString())
        If xmlSplitters Is Nothing Then
          xpath = New StringBuilder()
          xpath.Append("descendant::Amplifier[Handle='")
          xpath.Append(Handle)
          xpath.Append("']")
          xmlAmps = doc.Item("Amplifiers").SelectSingleNode(xpath.ToString())
          If xmlAmps Is Nothing Then
            xpath = New StringBuilder()
            xpath.Append("descendant::SignalPoint[Handle='")
            xpath.Append(Handle)
            xpath.Append("']")
            xmlSigPoints = doc.Item("SignalPoints").SelectSingleNode(xpath.ToString())
            If xmlSigPoints Is Nothing Then
              Return "Error"
            Else
              'Здесь необходимо реализовать привязку усилителя к узлам
              'Dim sx As String
              'sx = xmlSigPoints.Item("X").FirstChild.Value
              'Dim pt As Point3d = New Point3d(Double.Parse(xmlSigPoints.Item("X").FirstChild.Value.ToString), Double.Parse(xmlSigPoints.Item("Y").FirstChild.Value.ToString), Double.Parse(xmlSigPoints.Item("Z").FirstChild.Value.ToString))
              s = xmlSigPoints.Item("OpticalNode").FirstChild.Value
              Return s
            End If
          Else
            s = xmlAmps.Item("Handle").FirstChild.Value
            '  xmlAmps.RemoveAll()
            Return s
          End If
        Else
          s = xmlSplitters.GetAttribute("PARENT_CROSS")
          'xmlSplitters.RemoveAll()
          Return FindParentAmplifier(s, doc)
        End If
      Else
        s = xmlTaps.GetAttribute("PARENT_CROSS")
        'xmlTaps.RemoveAll()
        Return FindParentAmplifier(s, doc)
      End If
    Else
      s = xmlCabels.GetAttribute("PARENT_CROSS").ToString()
      'xmlCabels.RemoveAll()
      Return FindParentAmplifier(s, doc)
    End If
    Return Handle
  End Function

  Public Function FindParentAmplifier2(ByVal Handle As String, ByVal doc As XmlElement) As String
    Dim xpath As StringBuilder = New StringBuilder()
    Dim xmlNode As XmlElement

    Dim strXpath As String

    xpath.Append("descendant::Drawing//*[Handle='")
    xpath.Append(Handle)
    xpath.Append("']")

    strXpath = xpath.ToString()

    xmlNode = doc.SelectSingleNode(xpath.ToString())



    If xmlNode Is Nothing Then
      Return "Error:"
    Else

    End If

    Return "Ok"

  End Function



  Public Function FindOpticalNode(ByVal ed As Editor, ByVal pt As Point3d) As String

    Dim values() As TypedValue = {New TypedValue(DxfCode.BlockName, "AMP1_19,AMP1_20")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    Dim size As Long = 10
    Dim handle As String = ""


    Dim gripPts As Point3dCollection = New Point3dCollection()
    gripPts.Add(New Point3d(pt.X - size, pt.Y - size, pt.Z))
    gripPts.Add(New Point3d(pt.X + size, pt.Y - size, pt.Z))
    gripPts.Add(New Point3d(pt.X + size, pt.Y + size, pt.Z))
    gripPts.Add(New Point3d(pt.X - size, pt.Y + size, pt.Z))



    Res = ed.SelectCrossingPolygon(gripPts, sfilter)

    '    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("У данного источника нет узла.")
      ed.UpdateScreen()
      Return "NONE"
    End If

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()

    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      handle = blkRef.Handle.ToString()
    Next

    Return handle
  End Function

  Public Function ExportDeviceMarkers(ByVal ed As Editor) As XmlElement

    Dim xmDevTypes As XmlElement = xmdoc.CreateElement("DeviceTypes")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCDEVICELABEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт типов устройств.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об обозначениях оборудования.")
      ed.UpdateScreen()
      Return xmDevTypes
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmDeviceType As XmlElement = xmdoc.CreateElement("DeviceMarker")
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")

      Dim rb As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")
      If rb <> Nothing Then
        Dim rs As Array = rb.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          xmDeviceType.AppendChild(xmBezRef)
        Next
      End If
      Dim xmType As XmlElement = xmdoc.CreateElement("CAB_TYP") ' исправить на DEV_TYP
      Dim attcol As AttributeCollection = blkRef.AttributeCollection()
      Dim attid As ObjectId
      For Each attid In attcol
        Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
        If attref.Tag = "TYP" Then
          xmType.InnerText = attref.TextString
          xmDeviceType.AppendChild(xmType)
        End If
      Next
      xmX.InnerText = blkRef.Position.X.ToString()
      xmY.InnerText = blkRef.Position.Y.ToString()
      xmRotation.InnerText = blkRef.Rotation
      xmDeviceType.AppendChild(xmX)
      xmDeviceType.AppendChild(xmY)
      xmDeviceType.AppendChild(xmRotation)
      xmDevTypes.AppendChild(xmDeviceType)
    Next

    Return xmDevTypes
  End Function

  Public Function ExportLevelLines(ByVal ed As Editor) As XmlElement
    Dim PLCLevelLines As XmlElement = xmdoc.CreateElement("PLCLevelLines")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCLEVEL_TEXT_MARKLINE")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт линий уровней сигналов.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о линиях уровней сигналов.")
      ed.UpdateScreen()
      Return PLCLevelLines
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length
    Dim Id As ObjectId
    For Each Id In IdArray

      Dim plcLevelLine As XmlElement = xmdoc.CreateElement("PLCLevelLine")

      Dim xmpoint1 As XmlElement = xmdoc.CreateElement("Point")
      Dim xmpoint2 As XmlElement = xmdoc.CreateElement("Point")
      Dim StartPoint As StringBuilder = New StringBuilder()
      Dim EndPoint As StringBuilder = New StringBuilder()
      Dim ent As Entity = tm.GetObject(Id, OpenMode.ForRead)


      Select Case ent.GetType().Name.ToString
        Case "Line"
          Dim ln As Line = tm.GetObject(Id, OpenMode.ForRead)
          plcLevelLine.SetAttribute("handle", ln.Handle.ToString())
          StartPoint.Append(ln.StartPoint.X)
          StartPoint.Append(", ")
          StartPoint.Append(ln.StartPoint.Y)
          xmpoint1.InnerText = StartPoint.ToString()
          plcLevelLine.AppendChild(xmpoint1)
          EndPoint.Append(ln.EndPoint.X)
          EndPoint.Append(", ")
          EndPoint.Append(ln.EndPoint.Y)
          xmpoint2.InnerText = EndPoint.ToString()
          plcLevelLine.AppendChild(xmpoint2)
        Case "Polyline2d"
          Dim pline2d As Polyline2d = tm.GetObject(Id, OpenMode.ForRead)
          plcLevelLine.SetAttribute("handle", pline2d.Handle.ToString())
          StartPoint.Append(pline2d.GetPointAtParameter(0).X)
          StartPoint.Append(", ")
          StartPoint.Append(pline2d.GetPointAtParameter(0).Y)
          xmpoint1.InnerText = StartPoint.ToString()
          plcLevelLine.AppendChild(xmpoint1)
          EndPoint.Append(pline2d.GetPointAtParameter(1).X)
          EndPoint.Append(", ")
          EndPoint.Append(pline2d.GetPointAtParameter(1).Y)
          xmpoint2.InnerText = EndPoint.ToString()
          plcLevelLine.AppendChild(xmpoint2)
      End Select
      PLCLevelLines.AppendChild(plcLevelLine)
    Next

    Return PLCLevelLines
  End Function

  Public Function ExportLevelMarkers(ByVal ed As Editor) As XmlElement

    Dim xmLevelMarkers As XmlElement = xmdoc.CreateElement("LevelMarkers")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCLEVEL_TEXT_MARK")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт уровней сигналов.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об уровнях сигналов")
      ed.UpdateScreen()
      Return xmLevelMarkers
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmLevelMarker As XmlElement = xmdoc.CreateElement("LevelMarker")
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmAlign As XmlElement = xmdoc.CreateElement("Align")

      Select Case blkRef.Name
        Case "LEVTXT_H"
          xmAlign.InnerText = "Lower"
        Case "LEVTXT_L"
          xmAlign.InnerText = "Upper"
      End Select


      Dim rs2 As ResultBuffer = blkRef.GetXDataForApplication("BEZ_REF")

      If rs2 Is Nothing Then
        ed.WriteMessage("Проблема с экспортом уровней сигнала!!!!!!!!!!!!!!!!!!!!")
      Else
        Dim rs As Array = rs2.AsArray()
        For c As Integer = 1 To rs.Length - 1
          Dim xmBezRef As XmlElement = xmdoc.CreateElement("BEZ_REF")
          xmBezRef.InnerText = rs(c).Value
          xmLevelMarker.AppendChild(xmBezRef)
        Next
        Dim xmType As XmlElement = xmdoc.CreateElement("PLCLevel")
        Dim attcol As AttributeCollection = blkRef.AttributeCollection()
        Dim attid As ObjectId
        For Each attid In attcol
          Dim attref As AttributeReference = tm.GetObject(attid, OpenMode.ForRead, True)
          If attref.Tag = "PEGEL" Then
            xmType.InnerText = attref.TextString
            xmLevelMarker.AppendChild(xmType)
          End If
        Next
        xmX.InnerText = blkRef.Position.X.ToString()
        xmY.InnerText = blkRef.Position.Y.ToString()
        xmRotation.InnerText = blkRef.Rotation
        xmLevelMarker.AppendChild(xmX)
        xmLevelMarker.AppendChild(xmY)
        xmLevelMarker.AppendChild(xmRotation)
        xmLevelMarker.AppendChild(xmAlign)
        xmLevelMarkers.AppendChild(xmLevelMarker)
      End If
    Next
    Return xmLevelMarkers

  End Function


  Public Function ExportStreetNames(ByVal ed As Editor) As XmlElement

    Dim xmStreetNames As XmlElement = xmdoc.CreateElement("StreetNames")
    Dim values() As TypedValue = {New TypedValue(DxfCode.LayerName, "STREET_NAME"), _
    New TypedValue(DxfCode.Start, "TEXT"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт названий улиц.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о названиях улиц.")
      ed.UpdateScreen()
      Return xmStreetNames
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim StrName As DBText = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmStreetName As XmlElement = xmdoc.CreateElement("StreetName")
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmText As XmlElement = xmdoc.CreateElement("Text")
      xmX.InnerText = StrName.Position.X.ToString()
      xmY.InnerText = StrName.Position.Y.ToString()
      xmRotation.InnerText = StrName.Rotation
      xmText.InnerText = StrName.TextString
      xmStreetName.AppendChild(xmX)
      xmStreetName.AppendChild(xmY)
      xmStreetName.AppendChild(xmRotation)
      xmStreetName.AppendChild(xmText)
      xmStreetNames.AppendChild(xmStreetName)
    Next

    Return xmStreetNames

  End Function


  Public Function ExportTextStrings(ByVal ed As Editor) As XmlElement

    Dim xmTextStrings As XmlElement = xmdoc.CreateElement("TextStrings")
    Dim values() As TypedValue = {New TypedValue(DxfCode.LayerName, "TEXT"), _
    New TypedValue(DxfCode.Start, "TEXT"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт текстовой информации.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет текстовой информации")
      ed.UpdateScreen()
      Return xmTextStrings
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim Text As DBText = tm.GetObject(Id, OpenMode.ForRead)
      Dim xmTextString As XmlElement = xmdoc.CreateElement("TextString")
      Dim xmX As XmlElement = xmdoc.CreateElement("X")
      Dim xmY As XmlElement = xmdoc.CreateElement("Y")
      Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
      Dim xmText As XmlElement = xmdoc.CreateElement("Text")
      xmX.InnerText = Text.Position.X.ToString()
      xmY.InnerText = Text.Position.Y.ToString()
      xmRotation.InnerText = Text.Rotation
      xmText.InnerText = Text.TextString
      xmTextString.AppendChild(xmX)
      xmTextString.AppendChild(xmY)
      xmTextString.AppendChild(xmRotation)
      xmTextString.AppendChild(xmText)
      xmTextStrings.AppendChild(xmTextString)
    Next

    Return xmTextStrings

  End Function

  Public Function ExportFrontDoors(ByVal ed As Editor) As XmlElement

    Dim xmFrontDoors As XmlElement = xmdoc.CreateElement("FrontDoors")
    Dim values() As TypedValue = {New TypedValue(DxfCode.LayerName, "ПОДЪЕЗДЫ, Подъезды, подъезды"), _
    New TypedValue(DxfCode.Start, "TEXT,MTEXT"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт номеров подъездов.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о подъездах.")
      ed.UpdateScreen()
      Return xmFrontDoors
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim ent As Entity = tm.GetObject(Id, OpenMode.ForRead)
      Select Case ent.GetType().Name.ToString()
        Case "DBText"
          Dim Text As DBText = tm.GetObject(Id, OpenMode.ForRead)
          Dim xmFrontDoor As XmlElement = xmdoc.CreateElement("FrontDoor")
          Dim xmX As XmlElement = xmdoc.CreateElement("X")
          Dim xmY As XmlElement = xmdoc.CreateElement("Y")
          Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
          Dim xmText As XmlElement = xmdoc.CreateElement("Text")
          xmX.InnerText = Text.Position.X.ToString()
          xmY.InnerText = Text.Position.Y.ToString()
          xmRotation.InnerText = Text.Rotation
          xmText.InnerText = Text.TextString
          xmFrontDoor.AppendChild(xmX)
          xmFrontDoor.AppendChild(xmY)
          xmFrontDoor.AppendChild(xmRotation)
          xmFrontDoor.AppendChild(xmText)
          xmFrontDoors.AppendChild(xmFrontDoor)
        Case "MText"
          Dim Text As MText = tm.GetObject(Id, OpenMode.ForRead)
          Dim xmFrontDoor As XmlElement = xmdoc.CreateElement("FrontDoor")
          Dim xmX As XmlElement = xmdoc.CreateElement("X")
          Dim xmY As XmlElement = xmdoc.CreateElement("Y")
          Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
          Dim xmText As XmlElement = xmdoc.CreateElement("Text")
          xmX.InnerText = Text.Location.X.ToString()
          xmY.InnerText = Text.Location.Y.ToString()
          xmRotation.InnerText = Text.Rotation
          xmText.InnerText = Text.Text()
          xmFrontDoor.AppendChild(xmX)
          xmFrontDoor.AppendChild(xmY)
          xmFrontDoor.AppendChild(xmRotation)
          xmFrontDoor.AppendChild(xmText)
          xmFrontDoors.AppendChild(xmFrontDoor)
      End Select
    Next

    Return xmFrontDoors

  End Function


  Public Function ExportTextLabels(ByVal ed As Editor) As XmlElement

    Dim xmTextLabels As XmlElement = xmdoc.CreateElement("TextLabels")
    Dim values() As TypedValue = {New TypedValue(DxfCode.LayerName, "НАДПИСИ,НОМЕРА,Печать,Location_Frame*"), _
    New TypedValue(DxfCode.Start, "TEXT,MTEXT"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт различных надписей.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("Надписи не найдены.")
      ed.UpdateScreen()
      Return xmTextLabels
    End If
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId

    For Each Id In IdArray
      Dim ent As Entity = tm.GetObject(Id, OpenMode.ForRead)
      Select Case ent.GetType().Name.ToString()
        Case "DBText"
          Dim Text As DBText = tm.GetObject(Id, OpenMode.ForRead)
          Dim xmTextLabel As XmlElement = xmdoc.CreateElement("TextLabel")
          Dim xmX As XmlElement = xmdoc.CreateElement("X")
          Dim xmY As XmlElement = xmdoc.CreateElement("Y")
          Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
          Dim xmText As XmlElement = xmdoc.CreateElement("Text")
          xmX.InnerText = Text.Position.X.ToString()
          xmY.InnerText = Text.Position.Y.ToString()
          xmRotation.InnerText = Text.Rotation
          xmText.InnerText = Text.TextString
          xmTextLabel.AppendChild(xmX)
          xmTextLabel.AppendChild(xmY)
          xmTextLabel.AppendChild(xmRotation)
          xmTextLabel.AppendChild(xmText)
          xmTextLabels.AppendChild(xmTextLabel)
        Case "MText"
          Dim Text As MText = tm.GetObject(Id, OpenMode.ForRead)
          Dim xmTextLabel As XmlElement = xmdoc.CreateElement("TextLabel")
          Dim xmX As XmlElement = xmdoc.CreateElement("X")
          Dim xmY As XmlElement = xmdoc.CreateElement("Y")
          Dim xmRotation As XmlElement = xmdoc.CreateElement("Rotation")
          Dim xmText As XmlElement = xmdoc.CreateElement("Text")
          xmX.InnerText = Text.Location.X.ToString()
          xmY.InnerText = Text.Location.Y.ToString()
          xmRotation.InnerText = Text.Rotation
          xmText.InnerText = Text.Text()
          xmTextLabel.AppendChild(xmX)
          xmTextLabel.AppendChild(xmY)
          xmTextLabel.AppendChild(xmRotation)
          xmTextLabel.AppendChild(xmText)
          xmTextLabels.AppendChild(xmTextLabel)
        Case Else
      End Select
    Next
    Return xmTextLabels

  End Function


  ' Define command 'Exp'
  <CommandMethod("Exp")> _
  Public Sub Exp()
    ' Type your code here
    xmdoc = New XmlDocument

    Dim name As XmlElement = xmdoc.CreateElement("Name")

    'Создание xml документа
    Dim drawing As XmlElement = xmdoc.CreateElement("Drawing")
    '         Dim layer As XmlElement = xmdoc.CreateElement("Layer")
    ' Открытие базы данных чертежа
    Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database
    tm = db.TransactionManager
    Dim MyT As Transaction = tm.StartTransaction()

    Dim acaddoc As Document = Application.DocumentManager.MdiActiveDocument


    Dim ed As Editor = acaddoc.Editor

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт чертежа начат.")
    ed.UpdateScreen()

    'Отсюда необходимо выдрать имя файла
    'dsDrawing.DataSetName = db.Filename
    Dim xmDrawingName As XmlElement = xmdoc.CreateElement("DrawingName")
    xmDrawingName.InnerText = db.Filename.Substring(db.Filename.Length - 12)
    drawing.AppendChild(xmDrawingName)

    drawing.AppendChild(ExportSignalPoint(ed))
    drawing.AppendChild(ExportAddress(ed))
    drawing.AppendChild(ExportMap(ed))
    drawing.AppendChild(ExportCable(ed))
    drawing.AppendChild(ExportAmplifiers(ed))
    drawing.AppendChild(ExportTaps(ed))
    drawing.AppendChild(ExportSplitters(ed))
    drawing.AppendChild(ExportSignalConnections(ed))

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Построение дерева усилителей.")
    ed.UpdateScreen()
    drawing = Analyse(drawing)

    drawing.AppendChild(ExportOutlets(ed))
    drawing.AppendChild(ExportStreetNames(ed))
    drawing.AppendChild(ExportTextStrings(ed))
    drawing.AppendChild(ExportFrontDoors(ed))
    drawing.AppendChild(ExportTerminators(ed))
    drawing.AppendChild(ExportCableMarkers(ed))
    drawing.AppendChild(ExportDeviceMarkers(ed))
    drawing.AppendChild(ExportLevelLines(ed))
    drawing.AppendChild(ExportLevelMarkers(ed))
    drawing.AppendChild(ExportTextLabels(ed))
    drawing.AppendChild(ExportPowerSources(ed))

    xmdoc.AppendChild(drawing)
    '        Console.WriteLine(doc.OuterXml)
    Dim sfile As String
    sfile = db.Filename.Substring(db.Filename.Length - 12).Insert(12, ".xml")

    If Not Directory.Exists(db.Filename.Substring(0, db.Filename.Length - 12).Insert(db.Filename.Length - 12, "export_data")) Then
      Directory.CreateDirectory(db.Filename.Substring(0, db.Filename.Length - 12).Insert(db.Filename.Length - 12, "export_data"))
    End If
    Directory.SetCurrentDirectory(db.Filename.Substring(0, db.Filename.Length - 12).Insert(db.Filename.Length - 12, "export_data"))

    xmdoc.Save(sfile)
    xmdoc.RemoveAll()

    ed.WriteMessage(ControlChars.CrLf)
    ed.WriteMessage("Экспорт чертежа завершен.")
    ed.UpdateScreen()
    ' Type your code here

    MyT.Commit()
    MyT.Dispose()
    tm.Dispose()

  End Sub

  ' Define command 'Exp'
  <CommandMethod("TestProject")> _
  Public Sub TestProject()
    Dim count As Integer = 0
    Dim acaddoc As Document = Application.DocumentManager.MdiActiveDocument
    Dim ed As Editor = acaddoc.Editor
        ed.WriteMessage(vbLf & "Проверка проекта начата.")
    ed.UpdateScreen()

    CreateLayers()
    count += TestAddress(ed)
    count += TestAmplifiers(ed)
    count += TestTaps(ed)
    count += TestSplitters(ed)
    count += TestPowerSources(ed)
    ed.WriteMessage(vbLf & "Тест чертежа завершен.")
    ed.WriteMessage(vbLf & "В проекте " & count & " ошибок.")
    ed.UpdateScreen()
  End Sub

  Public Function TestAddress(ByVal ed As Editor)
    Dim count As Integer = 0
    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    'Dim xmlAddress As XmlElement = xmdoc.CreateElement("HouseNumbers")
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCHOUSENUMBER"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(vbLf & "Проверка номеров домов.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If Res.Status = PromptStatus.Error Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("Чертеж не содержит номеров домов.")
      ed.UpdateScreen()
      Return 0
    End If

    Dim db As Database = doc.Database
    Dim MyT As Transaction = db.TransactionManager.StartTransaction()
    Dim acBlkTbl As BlockTable
    acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    Dim acBlkTblRec As BlockTableRecord
    acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()

    Dim Id As ObjectId
    Dim c As Long = 0
    ReDim Address(SS.Count)

    For Each Id In IdArray
      Dim hn As BlockReference = MyT.GetObject(Id, OpenMode.ForRead)
      Dim rs As Array = hn.GetXDataForApplication("ESTATE_DATAS").AsArray()
      Address(c).Street = rs(1).Value
      Address(c).Number = rs(3).Value
      Address(c).AddInfo = rs(4).Value
      Address(c).ZipCode = rs(11).Value

      Dim rsFID As ResultBuffer = hn.GetXDataForApplication("FID")
      If rsFID = Nothing Then
        count += 1
        Dim acLine1 As Line = New Line(New Point3d(hn.Position.X - 5, hn.Position.Y - 5, 0), _
        New Point3d(hn.Position.X + 5, hn.Position.Y + 5, 0))
        Dim acLine2 As Line = New Line(New Point3d(hn.Position.X - 5, hn.Position.Y + 5, 0), _
        New Point3d(hn.Position.X + 5, hn.Position.Y - 5, 0))
        acLine1.Layer = "ErrorAddress"
        acLine2.Layer = "ErrorAddress"
        acLine1.SetDatabaseDefaults()
        acLine2.SetDatabaseDefaults()
        acBlkTblRec.AppendEntity(acLine1)
        MyT.AddNewlyCreatedDBObject(acLine1, True)
        acBlkTblRec.AppendEntity(acLine2)
        MyT.AddNewlyCreatedDBObject(acLine2, True)
        Address(c).Fid = 0
      Else
        Address(c).Fid = rsFID.AsArray(2).Value
      End If

      c = c + 1
    Next
    MyT.Commit()
    MyT.Dispose()

    Return count

  End Function

  Public Function TestAmplifiers(ByVal ed As Editor)
    Dim c As Integer = 0
    Dim tdoc As Document = Application.DocumentManager.MdiActiveDocument

    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXAMPLIFIER")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(vbLf & "Проверка усилителей.")
    ed.UpdateScreen()


    Res = ed.SelectAll(sfilter)

    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об усилителях")
      ed.UpdateScreen()
      Return 0
    End If

    Dim db As Database = tdoc.Database
    Dim MyT As Transaction = db.TransactionManager.StartTransaction()
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim acBlkTbl As BlockTable
    acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    Dim acBlkTblRec As BlockTableRecord
    acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim count As Integer
    count = IdArray.Length

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = MyT.GetObject(Id, OpenMode.ForRead)

      'Запись адреса усилителя
      Dim rsEstate As ResultBuffer = blkRef.GetXDataForApplication("ESTATE_DATAS")
      If rsEstate <> Nothing Then
        Dim rs1 As Array = rsEstate.AsArray()
        Select Case FindFid(rs1(3).Value, rs1(4).Value, rs1(1).Value, rs1(11).value)
          Case 0
            c += 1
            Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
            Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
            acLine1.Layer = "ErrorAddressDevice"
            acLine2.Layer = "ErrorAddressDevice"
            acLine1.SetDatabaseDefaults()
            acLine2.SetDatabaseDefaults()
            acBlkTblRec.AppendEntity(acLine1)
            MyT.AddNewlyCreatedDBObject(acLine1, True)
            acBlkTblRec.AppendEntity(acLine2)
            MyT.AddNewlyCreatedDBObject(acLine2, True)
          Case -1
            c += 1
            Dim acCircle As Circle = New Circle(New Point3d(blkRef.Position.X, _
            blkRef.Position.Y, 0), New Vector3d(0, 0, 1), 5)
            acCircle.SetDatabaseDefaults()
            acCircle.Layer = "ErrorAddressDevice"
            acBlkTblRec.AppendEntity(acCircle)
            MyT.AddNewlyCreatedDBObject(acCircle, True)
        End Select
        'xmAmplifier.AppendChild(xmEstate)
      Else
        Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
        Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
        acLine1.Layer = "ErrorAddressDevice"
        acLine2.Layer = "ErrorAddressDevice"
        acLine1.SetDatabaseDefaults()
        acLine2.SetDatabaseDefaults()
        acBlkTblRec.AppendEntity(acLine1)
        MyT.AddNewlyCreatedDBObject(acLine1, True)
        acBlkTblRec.AppendEntity(acLine2)
        MyT.AddNewlyCreatedDBObject(acLine2, True)
      End If
    Next
    MyT.Commit()
    MyT.Dispose()
    Return c
  End Function

  Public Function TestTaps(ByVal ed As Editor)
    Dim c As Integer = 0
    Dim tdoc As Document = Application.DocumentManager.MdiActiveDocument
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXTAP")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(vbLf & "Проверка ответвителей.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)
    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных об ответвителях")
      ed.UpdateScreen()
      Return 0
    End If

    Dim db As Database = tdoc.Database
    Dim MyT As Transaction = db.TransactionManager.StartTransaction()
    Dim acBlkTbl As BlockTable
    acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    Dim acBlkTblRec As BlockTableRecord
    acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = MyT.GetObject(Id, OpenMode.ForRead)

      'Запись адреса усилителя
      Dim rsEstate As ResultBuffer = blkRef.GetXDataForApplication("ESTATE_DATAS")

      If rsEstate <> Nothing Then
        Dim rs1 As Array = rsEstate.AsArray()
        Select Case FindFid(rs1(3).Value, rs1(4).Value, rs1(1).Value, rs1(11).value)
          Case 0
            c += 1
            Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
            Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
            acLine1.Layer = "ErrorAddressDevice"
            acLine2.Layer = "ErrorAddressDevice"
            acLine1.SetDatabaseDefaults()
            acLine2.SetDatabaseDefaults()
            acBlkTblRec.AppendEntity(acLine1)
            MyT.AddNewlyCreatedDBObject(acLine1, True)
            acBlkTblRec.AppendEntity(acLine2)
            MyT.AddNewlyCreatedDBObject(acLine2, True)
          Case -1
            c += 1
            Dim acCircle As Circle = New Circle(New Point3d(blkRef.Position.X, _
            blkRef.Position.Y, 0), New Vector3d(0, 0, 1), 5)
            acCircle.SetDatabaseDefaults()
            acCircle.Layer = "ErrorAddressDevice"
            acBlkTblRec.AppendEntity(acCircle)
            MyT.AddNewlyCreatedDBObject(acCircle, True)
        End Select
        'xmAmplifier.AppendChild(xmEstate)
      Else
        c += 1
        Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
        Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
        acLine1.Layer = "ErrorAddressDevice"
        acLine2.Layer = "ErrorAddressDevice"
        acLine1.SetDatabaseDefaults()
        acLine2.SetDatabaseDefaults()
        acBlkTblRec.AppendEntity(acLine1)
        MyT.AddNewlyCreatedDBObject(acLine1, True)
        acBlkTblRec.AppendEntity(acLine2)
        MyT.AddNewlyCreatedDBObject(acLine2, True)
      End If
      'xmAmplifiers.AppendChild(xmAmplifier)
    Next
    MyT.Commit()
    MyT.Dispose()
    Return c
  End Function

  Public Function TestSplitters(ByVal ed As Editor)
    Dim c As Integer = 0
    Dim tdoc As Document = Application.DocumentManager.MdiActiveDocument

    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCCOAXSPLITTER")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(vbLf & "Проверка делителей.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)

    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о делителях")
      ed.UpdateScreen()
      Return 0
    End If
    Dim db As Database = tdoc.Database
    Dim MyT As Transaction = db.TransactionManager.StartTransaction()

    Dim acBlkTbl As BlockTable
    acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    Dim acBlkTblRec As BlockTableRecord
    acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = MyT.GetObject(Id, OpenMode.ForRead)

      'Запись адреса усилителя
      Dim rsEstate As ResultBuffer = blkRef.GetXDataForApplication("ESTATE_DATAS")

      If rsEstate <> Nothing Then
        Dim rs1 As Array = rsEstate.AsArray()
        Select Case FindFid(rs1(3).Value, rs1(4).Value, rs1(1).Value, rs1(11).value)
          Case 0
            c += 1
            Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
            Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
            acLine1.Layer = "ErrorAddressDevice"
            acLine2.Layer = "ErrorAddressDevice"
            acLine1.SetDatabaseDefaults()
            acLine2.SetDatabaseDefaults()
            acBlkTblRec.AppendEntity(acLine1)
            MyT.AddNewlyCreatedDBObject(acLine1, True)
            acBlkTblRec.AppendEntity(acLine2)
            MyT.AddNewlyCreatedDBObject(acLine2, True)
          Case -1
            Dim acCircle As Circle = New Circle(New Point3d(blkRef.Position.X, _
            blkRef.Position.Y, 0), New Vector3d(0, 0, 1), 5)
            acCircle.SetDatabaseDefaults()
            acCircle.Layer = "ErrorAddressDevice"
            acBlkTblRec.AppendEntity(acCircle)
            MyT.AddNewlyCreatedDBObject(acCircle, True)
            c += 1
        End Select
        'xmAmplifier.AppendChild(xmEstate)
      Else
        c += 1
        Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
        Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
        acLine1.Layer = "ErrorAddressDevice"
        acLine2.Layer = "ErrorAddressDevice"
        acLine1.SetDatabaseDefaults()
        acLine2.SetDatabaseDefaults()
        acBlkTblRec.AppendEntity(acLine1)
        MyT.AddNewlyCreatedDBObject(acLine1, True)
        acBlkTblRec.AppendEntity(acLine2)
        MyT.AddNewlyCreatedDBObject(acLine2, True)
      End If
      'xmAmplifiers.AppendChild(xmAmplifier)
    Next
    MyT.Commit()
    MyT.Dispose()
    Return c
  End Function

  Public Function TestPowerSources(ByVal ed As Editor)
    Dim c As Integer = 0
    Dim tdoc As Document = Application.DocumentManager.MdiActiveDocument
    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCREMOTEPOWERSOURCE")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(vbLf & "Проверка блоков питания.")
    ed.UpdateScreen()

    Res = ed.SelectAll(sfilter)

    If (Res.Status = PromptStatus.Error) Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("В чертеже нет данных о блоках питания")
      ed.UpdateScreen()
      Return 0
    End If

    Dim db As Database = tdoc.Database
    Dim MyT As Transaction = db.TransactionManager.StartTransaction()
    Dim acBlkTbl As BlockTable
    acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    Dim acBlkTblRec As BlockTableRecord
    acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value

    Dim IdArray As ObjectId() = SS.GetObjectIds()

    Dim Id As ObjectId
    For Each Id In IdArray
      Dim blkRef As BlockReference = MyT.GetObject(Id, OpenMode.ForRead)

      'Запись адреса усилителя
      Dim rbEstate As ResultBuffer = blkRef.GetXDataForApplication("ESTATE_DATAS")

      If rbEstate <> Nothing Then
        Dim arEstate As Array = blkRef.GetXDataForApplication("ESTATE_DATAS").AsArray()

        Select Case FindFid(arEstate(3).Value, arEstate(4).Value, arEstate(1).Value, arEstate(11).value)
          Case 0
            c += 1
            Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
            Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
            New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
            acLine1.Layer = "ErrorAddressDevice"
            acLine2.Layer = "ErrorAddressDevice"
            acLine1.SetDatabaseDefaults()
            acLine2.SetDatabaseDefaults()
            acBlkTblRec.AppendEntity(acLine1)
            MyT.AddNewlyCreatedDBObject(acLine1, True)
            acBlkTblRec.AppendEntity(acLine2)
            MyT.AddNewlyCreatedDBObject(acLine2, True)
          Case -1
            c += 1
            Dim acCircle As Circle = New Circle(New Point3d(blkRef.Position.X, _
            blkRef.Position.Y, 0), New Vector3d(0, 0, 1), 5)
            acCircle.SetDatabaseDefaults()
            acCircle.Layer = "ErrorAddressDevice"
            acBlkTblRec.AppendEntity(acCircle)
            MyT.AddNewlyCreatedDBObject(acCircle, True)
        End Select
        'xmAmplifier.AppendChild(xmEstate)
      Else
        c += 1
        Dim acLine1 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y - 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y + 5, 0))
        Dim acLine2 As Line = New Line(New Point3d(blkRef.Position.X - 5, blkRef.Position.Y + 5, 0), _
        New Point3d(blkRef.Position.X + 5, blkRef.Position.Y - 5, 0))
        acLine1.SetDatabaseDefaults()
        acLine2.SetDatabaseDefaults()
        acLine1.Layer = "ErrorAddressDevice"
        acLine2.Layer = "ErrorAddressDevice"
        acBlkTblRec.AppendEntity(acLine1)
        MyT.AddNewlyCreatedDBObject(acLine1, True)
        acBlkTblRec.AppendEntity(acLine2)
        MyT.AddNewlyCreatedDBObject(acLine2, True)
      End If
      'xmAmplifiers.AppendChild(xmAmplifier)
    Next
    MyT.Commit()
    MyT.Dispose()
    Return c
  End Function

  Public Sub CreateLayers()
    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    Dim db As Database = doc.Database
    Dim MyT As Transaction = db.TransactionManager.StartTransaction()

    Dim AcLayerTbl As LayerTable
    AcLayerTbl = MyT.GetObject(db.LayerTableId, OpenMode.ForRead)
    Dim sLayerName As String = "ErrorAddress"
    If AcLayerTbl.Has(sLayerName) = False Then
      Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()
      acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1)
      acLyrTblRec.Name = sLayerName
      AcLayerTbl.UpgradeOpen()
      AcLayerTbl.Add(acLyrTblRec)
      MyT.AddNewlyCreatedDBObject(acLyrTblRec, True)
    End If
    sLayerName = "ErrorAddressDevice"
    If AcLayerTbl.Has(sLayerName) = False Then
      Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()
      acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1)
      acLyrTblRec.Name = sLayerName
      AcLayerTbl.UpgradeOpen()
      AcLayerTbl.Add(acLyrTblRec)
      MyT.AddNewlyCreatedDBObject(acLyrTblRec, True)
    End If
    MyT.Commit()
    MyT.Dispose()
  End Sub

  <CommandMethod("UnitCount")> _
  Public Sub UnitCount()
    Dim acaddoc As Document = Application.DocumentManager.MdiActiveDocument
    Dim ed As Editor = acaddoc.Editor
    '    GetAddressList(ed)
    ReDim Address(15)
    GetAddressList(ed)
  End Sub

  Public Sub GetAddressList(ByVal ed As Editor)
    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    Dim db As Database = doc.Database

    Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCHOUSENUMBER"), _
    New TypedValue(DxfCode.LayoutName, "MODEL")}
    Dim sfilter As New SelectionFilter(values)
    Dim Res As PromptSelectionResult
    ed.WriteMessage(vbLf & "Выберите дома для нового кластера.")
    ed.UpdateScreen()


    Res = ed.GetSelection(sfilter)
    'End Select

    If Res.Status <> PromptStatus.OK Then
      ed.WriteMessage(ControlChars.CrLf)
      ed.WriteMessage("Ваш выбор не содержит номеров домов.")
      ed.UpdateScreen()
      Return
    End If
    Using MyT As Transaction = db.TransactionManager.StartTransaction()
      Dim acBlkTbl As BlockTable
      acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
      'Dim acBlkTblRec As BlockTableRecord
      'acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    End Using


    Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    Dim IdArray As ObjectId() = SS.GetObjectIds()
    Dim Id As ObjectId
    Dim c As Long = 0
    ReDim ad(SS.Count - 1)
    Using MyT As Transaction = db.TransactionManager.StartTransaction()
      For Each Id In IdArray

        Dim hn As BlockReference = MyT.GetObject(Id, OpenMode.ForRead)

        Dim rs As Array = hn.GetXDataForApplication("ESTATE_DATAS").AsArray()
        Dim rsFID As ResultBuffer = hn.GetXDataForApplication("FID")

        If rsFID = Nothing Then

        Else
          ad(c).Fid = rsFID.AsArray(2).Value
          ad(c).Street = rs(1).ToString()
          ad(c).Number = rs(3).ToString()
        End If
        c = c + 1

      Next
    End Using
    ad = GetExcelData(ad)

    Dim AdData As AddressDatas
    Dim quartes As Integer
    quartes = 0
    For Each AdData In ad
      If Not AdData.Gpon Then quartes = quartes + AdData.Quarters
    Next

    Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
    With pKeyOpts
      .Message = vbLf & "Выберите способ вставки данных по квартирам"
      .Keywords.Add("Текст")
      .Keywords.Add("таБлица")
      '.Keywords.Add("таблица с Выгрузкой в Excel")
      .AllowNone = False
    End With
    Dim pKeyRes As PromptResult = ed.GetKeywords(pKeyOpts)
    Dim s As StringBuilder = New StringBuilder()
    s.Append("В выбраных домах количество квартир равняется ")
    s.Append(quartes)
    ed.WriteMessage(s.ToString())
    ed.WriteMessage(ControlChars.CrLf)
    ed.UpdateScreen()

    Select Case pKeyRes.StringResult
      Case "Текст"
        Dim txt As DBText = New DBText()
        txt.TextString = quartes.ToString & "кв."
        txt.Height = 2.5
        Using MyT As Transaction = db.TransactionManager.StartTransaction()
          Dim acBlkTbl As BlockTable
          acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
          Dim acBlkTblRec As BlockTableRecord
          acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

          'Dim s As StringBuilder = New StringBuilder()
          's.Append(address.Street)
          's.Append(", д.")
          's.Append(address.Number)
          Dim pt As Point3d

          Dim ptres As PromptPointResult
          ptres = ed.GetPoint("Выберите точку вставки текста")
          pt = New Point3d(ptres.Value.X, ptres.Value.Y, ptres.Value.Z)
          txt.Position = pt
          '          tbl.GenerateLayout()
          acBlkTblRec.AppendEntity(txt)
          MyT.AddNewlyCreatedDBObject(txt, True)
          MyT.Commit()
          MyT.Dispose()
        End Using
      Case "таБлица"
        BuildAddressTable(ed, ad)

        'Case "таблица с Выгрузкой в Excel"
        '  BuildAddressTable(ed, ad)
    End Select

  End Sub

  Public Function GetExcelData(ByVal Ad As Array) As Array
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim rngRes As Excel.Range

    Dim l As Integer
    l = Ad.Length

    xlApp = CreateObject("Excel.Application")
    xlBook = xlApp.Workbooks.Open("C:\\Petrograd.xlsx")
    xlSheet = xlBook.Worksheets(2)

    xlRange = xlSheet.Range("$A$1").CurrentRegion
    Dim house As AddressDatas
    Dim c As Integer
    c = 0
    For Each house In Ad
      If house.Fid <> 0 Then
        rngRes = xlRange.Resize(xlRange.Rows.Count, 1).Find(What:=house.Fid, LookAt:=Excel.XlLookAt.xlWhole)
        If rngRes Is Nothing Then
          house.AddInfo = "Дома с данным objid нет в базе"
        Else
          house.Street = xlSheet.Range(rngRes.Address).Offset(0, 2).Value
          house.Number = xlSheet.Range(rngRes.Address).Offset(0, 3).Value
          house.AddInfo = xlSheet.Range(rngRes.Address).Offset(0, 4).Value
          house.Quarters = xlSheet.Range(rngRes.Address).Offset(0, 9).Value
          house.Clients = xlSheet.Range(rngRes.Address).Offset(0, 9).Value
          house.PDAll = xlSheet.Range(rngRes.Address).Offset(0, 10).Value
          house.PDActive = xlSheet.Range(rngRes.Address).Offset(0, 10).Value
          house.Fid = house.Fid

          house.Gpon = False
        End If


        'If xlSheet.Range(rngRes.Address).Offset(0, 36).Value = "есть" Then
        '  house.Gpon = True
        '  Dim s As StringBuilder = New StringBuilder()
        '  s.Append("В доме ")
        '  s.Append(house.Street)
        '  s.Append(" ,д.")
        '  s.Append(house.Number)
        '  s.Append(" имеется gpon сеть.")
        '  MsgBox(s.ToString())
        'Else
        '  house.Gpon = False
        'End If




        '   house.Gpon = xlSheet.Range(rngRes.Address).Offset(0, 36).Value
        '      house.TKT = True
      End If
        Ad(c) = house
        c = c + 1
    Next

    '    Array.Sort(Ad)

    xlRange = Nothing
    rngRes = Nothing

    xlBook.Close()

    xlApp.Quit()

    GC.Collect()

    Return Ad

  End Function

  Public Sub BuildAddressTable(ByVal ed As Editor, ByVal ad As Array)
    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    Dim db As Database = doc.Database
    Dim MyT As Transaction = db.TransactionManager.StartTransaction()
    Dim acBlkTbl As BlockTable
    acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    Dim acBlkTblRec As BlockTableRecord
    acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    Dim rows As Integer = ad.Length

    Dim tbl As Table
    tbl = New Table()
    tbl.TableStyle = db.Tablestyle
    tbl.InsertColumns(0, 9, 8)
    'tbl.InsertColumns(1, 50, 1)
    'tbl.InsertColumns(2, 10, 2)
    'tbl.InsertColumns(5, 15, 1)
    'tbl.InsertColumns(6, 15, 2)

    tbl.SetColumnWidth(0, 9)
    tbl.SetColumnWidth(1, 70)
    tbl.SetColumnWidth(2, 15)
    tbl.SetColumnWidth(3, 15)
    tbl.SetColumnWidth(4, 15)
    tbl.SetColumnWidth(5, 20)
    tbl.SetColumnWidth(6, 70)

    tbl.InsertRows(0, 7, 1)
    tbl.InsertRows(1, 4, rows + 1)

    tbl.SetValue(0, 0, "Поз.")
    tbl.SetValue(0, 1, "Адрес")
    tbl.SetValue(0, 2, "Абонентов")
    tbl.SetValue(0, 3, "ПД_Абоненты")
    '    tbl.SetValue(0, 4, "ПД_Актив")
    tbl.SetValue(0, 4, "Obj_id")
    tbl.SetValue(0, 5, "Drawing")
    tbl.SetValue(0, 6, "FullName")

    Dim row As Integer
    Dim address As AddressDatas
    Dim pdAll As Integer = 0
    Dim pdActive As Integer = 0
    Dim quarts As Integer = 0
    Dim nonPonQuarts As Integer = 0
    row = 1
    For Each address In ad
      Dim s As StringBuilder = New StringBuilder()
      If address.Fid = 0 Then
        s.Append("Нет адреса в базе")
      Else
        s.Append(address.Street.ToLower)
        s.Append(", д.")
        s.Append(address.Number)
        If (address.AddInfo <> "") Then
          s.Append(" к.")
          s.Append(address.AddInfo)
        End If
      End If
      tbl.SetValue(row, 0, row)
      tbl.SetValue(row, 1, s.ToString())
      tbl.SetValue(row, 2, address.Quarters)
      tbl.SetValue(row, 3, address.PDAll)
      '      tbl.SetValue(row, 4, address.PDActive)
      tbl.SetValue(row, 4, address.Fid.ToString)
      tbl.SetValue(row, 5, db.Filename.Substring(db.Filename.Length - 12))
      s.Append("_")
      s.Append(address.Fid.ToString)
      tbl.SetValue(row, 6, s.ToString)
      quarts = quarts + address.Quarters
      pdAll = pdAll + address.PDAll
      pdActive = pdActive + address.PDActive
      If Not address.Gpon Then nonPonQuarts = nonPonQuarts + address.Quarters
      row = row + 1
    Next
    tbl.SetValue(row, 2, quarts)
    tbl.SetValue(row, 3, pdAll)


    'tbl.SetValue(row, 4, pdActive)
    '    tbl.SetValue(row, 1, nonPonQuarts)



    Dim pt As Point3d

    Dim ptres As PromptPointResult
    ptres = ed.GetPoint("Выберите точку вставки таблицы:")
    pt = New Point3d(ptres.Value.X, ptres.Value.Y, ptres.Value.Z)
    tbl.Position = pt
    tbl.GenerateLayout()
    acBlkTblRec.AppendEntity(tbl)
    MyT.AddNewlyCreatedDBObject(tbl, True)
    MyT.Commit()
    MyT.Dispose()
  End Sub


  <CommandMethod("TableGen")> _
    Public Sub TableGen()
    Dim acaddoc As Document = Application.DocumentManager.MdiActiveDocument
    Dim ed As Editor = acaddoc.Editor
    '    GetAddressList(ed)
    ReDim Address(15)
    GetAddressList(ed)
  End Sub

  Public Sub GetTableList(ByVal ed As Editor)

    'Dim PLCLevelLines As XmlElement = xmdoc.CreateElement("PLCLevelLines")
    'Dim values() As TypedValue = {New TypedValue(DxfCode.ExtendedDataRegAppName, "PLCLEVEL_TEXT_MARKLINE")}
    'Dim sfilter As New SelectionFilter(values)
    'Dim Res As PromptSelectionResult

    'ed.WriteMessage(ControlChars.CrLf)
    'ed.WriteMessage("Экспорт линий уровней сигналов.")
    'ed.UpdateScreen()

    'Res = ed.SelectAll(sfilter)
    'If (Res.Status = PromptStatus.Error) Then
    '  ed.WriteMessage(ControlChars.CrLf)
    '  ed.WriteMessage("В чертеже нет данных о линиях уровней сигналов.")
    '  ed.UpdateScreen()
    '  Return 'PLCLevelLines
    'End If
    'Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value



    'If Res.Status <> PromptStatus.OK Then
    '  ed.WriteMessage(ControlChars.CrLf)
    '  ed.WriteMessage("Ваш выбор не содержит номеров домов.")
    '  ed.UpdateScreen()
    '  Return
    'End If

    'Using MyT As Transaction = db.TransactionManager.StartTransaction()
    '  Dim acBlkTbl As BlockTable
    '  acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    '  'Dim acBlkTblRec As BlockTableRecord
    '  'acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    'End Using


    'Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = Res.Value
    'Dim IdArray As ObjectId() = SS.GetObjectIds()
    'Dim Id As ObjectId
    'Dim c As Long = 0
    'ReDim ad(SS.Count - 1)
    'Using MyT As Transaction = db.TransactionManager.StartTransaction()
    '  For Each Id In IdArray

    '    Dim hn As BlockReference = MyT.GetObject(Id, OpenMode.ForRead)

    '    Dim rs As Array = hn.GetXDataForApplication("ESTATE_DATAS").AsArray()
    '    Dim rsFID As ResultBuffer = hn.GetXDataForApplication("FID")

    '    If rsFID = Nothing Then

    '    Else
    '      ad(c).Fid = rsFID.AsArray(2).Value
    '    End If
    '    c = c + 1

    '  Next
    'End Using
    'ad = GetExcelData(ad)

    'Dim AdData As AddressDatas
    'Dim quartes As Integer
    'quartes = 0
    'For Each AdData In ad
    '  If Not AdData.Gpon Then quartes = quartes + AdData.Quarters
    'Next

    'Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
    'With pKeyOpts
    '  .Message = vbLf & "Выберите способ вставки данных по квартирам"
    '  .Keywords.Add("Текст")
    '  .Keywords.Add("таБлица")
    '  .AllowNone = False
    'End With
    'Dim pKeyRes As PromptResult = ed.GetKeywords(pKeyOpts)
    'Dim s As StringBuilder = New StringBuilder()
    's.Append("В выбраных домах количество квартир равняется ")
    's.Append(quartes)
    'ed.WriteMessage(s.ToString())
    'ed.WriteMessage(ControlChars.CrLf)
    'ed.UpdateScreen()

    'Select Case pKeyRes.StringResult
    '  Case "Текст"
    '    Dim txt As DBText = New DBText()
    '    txt.TextString = quartes.ToString & "кв."
    '    txt.Height = 2.5
    '    Using MyT As Transaction = db.TransactionManager.StartTransaction()
    '      Dim acBlkTbl As BlockTable
    '      acBlkTbl = MyT.GetObject(db.BlockTableId, OpenMode.ForRead)
    '      Dim acBlkTblRec As BlockTableRecord
    '      acBlkTblRec = MyT.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

    '      'Dim s As StringBuilder = New StringBuilder()
    '      's.Append(address.Street)
    '      's.Append(", д.")
    '      's.Append(address.Number)
    '      Dim pt As Point3d

    '      Dim ptres As PromptPointResult
    '      ptres = ed.GetPoint("Выберите точку вставки текста")
    '      pt = New Point3d(ptres.Value.X, ptres.Value.Y, ptres.Value.Z)
    '      txt.Position = pt
    '      '          tbl.GenerateLayout()
    '      acBlkTblRec.AppendEntity(txt)
    '      MyT.AddNewlyCreatedDBObject(txt, True)
    '      MyT.Commit()
    '      MyT.Dispose()
    '    End Using
    '  Case "таБлица"
    '    BuildAddressTable(ed, ad)
    'End Select

  End Sub







  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
End Class