VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: MapTip
'
' Call Initialize in Form_Load to provide a Map, Timer,
' PictureBox, and Label. The Label control should be inside
' the PictureBox. The PictureBox's Appearance should be
' set to 0-Flat at design time because it can not be set
' at run time.
'
' Use SetLayer to make the MapTip work with a particular
' MapLayer and field name.
'
' Wire the MapTip to your form:
'  -Call Timer from the Timer's Timer event.
'  -Call MouseMove from the Map's MouseMove event.
'
Option Explicit

Private m_x As Single      ' current x position
Private m_y As Single      ' current y position
Private m_lastX As Single  ' x position when timer starts
Private m_lastY As Single  ' y position when timer starts

Private m_map As Object
Private m_timer As Timer
Private m_picture As PictureBox
Private m_label As Label

Private m_layer As MapLayer  ' layer to search
Private m_field As String    ' field to get ToolTip text from

Private Function DoSearch() As MapObjects2.Recordset
  Dim recs As MapObjects2.Recordset
  Dim pt As MapObjects2.Point
  Set pt = m_map.ToMapPoint(m_x, m_y)
  If m_layer.shapeType = moPolygon Then
    Set recs = m_layer.SearchShape(pt, moPointInPolygon, "")
  Else
    Set recs = m_layer.SearchByDistance(pt, m_map.ToMapDistance(100), "")
  End If
  Set DoSearch = recs
End Function

Public Sub Initialize(map As Object, tmr As Timer, pic As PictureBox, lbl As Label)
  Set m_map = map
  Set m_timer = tmr
  Set m_picture = pic
  Set m_label = lbl
  
  m_picture.Visible = False
  m_picture.BackColor = vbInfoBackground
  
  m_label.ForeColor = vbInfoText
  m_label.AutoSize = True
  m_label.BackStyle = 0 ' transparent
End Sub

Public Sub MouseMove(x As Single, y As Single)
  m_x = x
  m_y = y
  If m_timer.Interval = 0 Then  ' start the timer
    m_lastX = x
    m_lastY = y
    m_timer.Interval = 100
  Else  ' hide the tooltip
    m_picture.Visible = False
  End If
End Sub

Public Sub SetLayer(layer As MapLayer, fld As String)
  Set m_layer = layer
  m_field = fld
End Sub

Private Sub ShowTipText(text As String)
  'set the caption
  m_label.Caption = text
  m_label.Left = 50
  m_label.Top = 0
  
  ' position the picture
  m_picture.Left = m_map.Left + m_x
  m_picture.Top = m_map.Top + m_y + 290
  m_picture.Width = m_label.Width + 100
  m_picture.Visible = True
End Sub

Public Sub Timer()
  If m_x = m_lastX And m_y = m_lastY Then
    ' mouse didn't move
    m_timer.Interval = 0
    Dim recs As MapObjects2.Recordset
    Set recs = DoSearch
    If recs.EOF Then
      ' nothing at this location
      m_picture.Visible = False
    Else
      ' show the toolTip
      ShowTipText recs(m_field).Value
    End If
  Else ' start over at the current location
    m_lastX = m_x
    m_lastY = m_y
  End If
End Sub


