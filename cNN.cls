VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'Version 1.0.2

Private Type cDendrite
  ConnectionStrength  As Double
  ConnectionTarget    As Variant 'Expected array: 0=Layer, 1=Neuron
End Type

Private Type cNeuron
  Dendrites()         As cDendrite
  DendriteCount       As Long
  Value               As Double
  BaseValue           As Double
  Error               As Double
  LearningCoefficient As Double
End Type

Private Type cLayer
  Neurons()           As cNeuron
  NeuronCount         As Long
  AdjacentLayers      As Variant 'Expected array: 0=Layer
End Type

Private Type cMemoryBank
  Pointer             As cNeuron 'Value points to current slot
  Neurons()           As cNeuron
  NeuronCount         As Long
  'Slots()             As Double
  'SlotCount           As Long
End Type


Private Const CONST_LearningCoefficient = 1.5

Private Layers()      As cLayer
Private LayerCount    As Long
Private MemBanks()    As cMemoryBank
Private MemBankCount  As Long




Private Sub Class_Initialize()
  ReDim Layers(0)
  LayerCount = 0
  ReDim MemBanks(0)
  MemBankCount = 0
End Sub

Private Sub Class_Terminate()
  ReDim Layers(0)
  ReDim MemBanks(0)
End Sub



Public Sub InitializeNN(Structure As Variant, Optional Memory As Variant = Null)
  Dim i        As Long
  Dim j        As Long
  Dim k        As Long
  Dim x        As Variant
  Dim mLayer   As cLayer
  Dim mMemBank As cMemoryBank
  
  
  'Create Layers
  For i = 0 To UBound(Structure)
    mLayer.AdjacentLayers = Array(i - 1)
    ReDim mLayer.Neurons(1 To Structure(i))
    mLayer.NeuronCount = Structure(i)
    For j = 1 To mLayer.NeuronCount
      mLayer.Neurons(j).LearningCoefficient = CONST_LearningCoefficient
      mLayer.Neurons(j).BaseValue = GetRand
      mLayer.Neurons(j).Value = 0
      If i Then 'Not the first layer
        ReDim mLayer.Neurons(j).Dendrites(1 To Structure(i - 1))
        mLayer.Neurons(j).DendriteCount = Structure(i - 1)
        For k = 1 To Structure(i - 1)
          mLayer.Neurons(j).Dendrites(k).ConnectionStrength = GetRand
          mLayer.Neurons(j).Dendrites(k).ConnectionTarget = Array(1, i, k)
        Next
      End If
    Next
    AddLayer mLayer
  Next
  'Create Memory Banks
  If Not IsNull(Memory) Then
    For i = 0 To UBound(Memory)
      ReDim mMemBank.Neurons(1 To Memory(i)(1))
      mMemBank.NeuronCount = Memory(i)(1)
      'mMemBank.SlotCount = Memory(i)(0)
      'ReDim mMemBank.Slots(0, mMemBank.NeuronCount)
      'mMemBank.Pointer ********** Pointer init code goes here
      For j = 1 To mMemBank.NeuronCount
        mMemBank.Neurons(j).LearningCoefficient = CONST_LearningCoefficient
        mMemBank.Neurons(j).BaseValue = GetRand
        mMemBank.Neurons(j).Value = 0
        ReDim mMemBank.Neurons(j).Dendrites(1 To Structure(UBound(Structure) - 1))
        mMemBank.Neurons(j).DendriteCount = Structure(UBound(Structure) - 1)
        For k = 1 To Structure(UBound(Structure) - 1)
          mMemBank.Neurons(j).Dendrites(k).ConnectionStrength = GetRand
          mMemBank.Neurons(j).Dendrites(k).ConnectionTarget = Array(1, UBound(Structure), k)
        Next
      Next
      AddMemBank mMemBank
    Next
    'Connect second layer to Memory Banks
    If LayerCount > 1 Then
      For i = 1 To Layers(2).NeuronCount
        For j = 1 To MemBankCount
          For k = 1 To MemBanks(j).NeuronCount
            Layers(2).Neurons(i).DendriteCount = Layers(2).Neurons(i).DendriteCount + 1
            ReDim Preserve Layers(2).Neurons(i).Dendrites(1 To Layers(2).Neurons(i).DendriteCount)
            Layers(2).Neurons(i).Dendrites(Layers(2).Neurons(i).DendriteCount).ConnectionStrength = GetRand
            Layers(2).Neurons(i).Dendrites(Layers(2).Neurons(i).DendriteCount).ConnectionTarget = Array(2, j, k)
          Next
        Next
      Next
    End If
  End If
End Sub

Private Sub AddLayer(Layer As cLayer)
  LayerCount = LayerCount + 1
  ReDim Preserve Layers(0 To LayerCount)
  Layers(LayerCount) = Layer
End Sub

Private Sub AddMemBank(MemBank As cMemoryBank)
  MemBankCount = MemBankCount + 1
  ReDim Preserve MemBanks(0 To MemBankCount)
  MemBanks(MemBankCount) = MemBank
End Sub



Public Sub SetInput(Data As Variant)
  Dim i As Long
  
  If LayerCount Then
    For i = 1 To Layers(1).NeuronCount
      Layers(1).Neurons(i).Value = CDbl(Data(i - 1))
    Next
  End If
  Call Refresh
End Sub

Public Sub Refresh()
  Dim a As Variant, b As Variant
  Dim i As Long, j As Long, k As Long
  Dim x As Variant
  
  For i = 2 To LayerCount
    For j = 1 To Layers(i).NeuronCount
      DoEvents
      For k = 1 To Layers(i).Neurons(j).DendriteCount
        Select Case Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1
              x = Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case 2
              x = MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case Else
              x = 0
        End Select
        Layers(i).Neurons(j).Value = Layers(i).Neurons(j).Value + x * Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
      Next
      Layers(i).Neurons(j).Value = TransferFunction(Layers(i).Neurons(j).Value + Layers(i).Neurons(j).BaseValue)
    Next
  Next
  For i = 1 To MemBankCount
    For j = 1 To MemBanks(i).NeuronCount
      DoEvents
      For k = 1 To MemBanks(i).Neurons(j).DendriteCount
        Select Case MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1
              x = Layers(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case 2
              x = MemBanks(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case Else
              x = 0
        End Select
        MemBanks(i).Neurons(j).Value = MemBanks(i).Neurons(j).Value + x * MemBanks(i).Neurons(j).Dendrites(k).ConnectionStrength
      Next
      MemBanks(i).Neurons(j).Value = TransferFunction(MemBanks(i).Neurons(j).Value + MemBanks(i).Neurons(j).BaseValue)
    Next
    '**********Code for slot save and pointer adjustment goes here
  Next
End Sub

Public Function GetOutput(Optional Index As Long = 0) As Variant
  'return specific output neuron or entire output level
  Dim i As Long
  Dim x As Variant
  
  If Index Then
    GetOutput = CDbl(Layers(LayerCount).Neurons(Index).Value)
  Else
    ReDim Preserve x(1 To Layers(LayerCount).NeuronCount)
    For i = 1 To Layers(LayerCount).NeuronCount
      x(i) = CDbl(Layers(LayerCount).Neurons(i).Value)
    Next
    GetOutput = x
  End If
End Function

Public Sub Train(Target As Variant)
  AdjustWeights Target
End Sub



Private Sub AdjustWeights(Target As Variant)
  Dim i As Long, j As Long, k As Long
  Dim x As Variant
  
  'Correct Output Layer
  For i = 1 To Layers(LayerCount).NeuronCount
    Layers(LayerCount).Neurons(i).Error = CDbl(Target(i - 1)) - Layers(LayerCount).Neurons(i).Value 'Raw Error
  Next
  
  'Back Propigate
  For i = LayerCount To 2 Step -1
    For j = 1 To Layers(i).NeuronCount
      DoEvents
      Layers(i).Neurons(j).Error = Layers(i).Neurons(j).Error * Layers(i).Neurons(j).Value * (1 - Layers(i).Neurons(j).Value)
      For k = 1 To Layers(i).Neurons(j).DendriteCount
        Select Case Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1
              Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error = Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error + Layers(i).Neurons(j).Error * Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
          Case 2
              MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error = MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error + Layers(i).Neurons(j).Error * Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
          Case Else
              'Ignore
        End Select
      Next
    Next
  Next
  
  'Update Layer Weights
  For i = 2 To LayerCount
    For j = 1 To Layers(i).NeuronCount
      DoEvents
      For k = 1 To Layers(i).Neurons(j).DendriteCount
        Select Case Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1
              x = Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case 2
              x = MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case Else
              x = 0
        End Select
        Layers(i).Neurons(j).Dendrites(k).ConnectionStrength = Layers(i).Neurons(j).Dendrites(k).ConnectionStrength + Layers(i).Neurons(j).LearningCoefficient * x * Layers(i).Neurons(j).Error
      Next
      Layers(i).Neurons(j).BaseValue = Layers(i).Neurons(j).BaseValue + Layers(i).Neurons(j).LearningCoefficient * Layers(i).Neurons(j).Error
      Layers(i).Neurons(j).Error = 0
    Next
  Next
  'Update MemBank Weights
  For i = 2 To MemBankCount
    For j = 1 To MemBanks(i).NeuronCount
      DoEvents
      For k = 1 To MemBanks(i).Neurons(j).DendriteCount
        Select Case MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1
              x = Layers(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case 2
              x = MemBanks(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case Else
              x = 0
        End Select
        MemBanks(i).Neurons(j).Dendrites(k).ConnectionStrength = MemBanks(i).Neurons(j).Dendrites(k).ConnectionStrength + MemBanks(i).Neurons(j).LearningCoefficient * x * MemBanks(i).Neurons(j).Error
      Next
      MemBanks(i).Neurons(j).BaseValue = MemBanks(i).Neurons(j).BaseValue + Layers(i).Neurons(j).LearningCoefficient * MemBanks(i).Neurons(j).Error
      MemBanks(i).Neurons(j).Error = 0
    Next
  Next
End Sub


Private Function TransferFunction(InVal As Double) As Double
  'Uses Sigmoid Function to squash InVal into interval 0..1
  On Error Resume Next
  TransferFunction = 1 / (1 + Exp(-InVal))
  If Err Then
    TransferFunction = 0
    Err.Clear
  End If
End Function

Private Function GetRand() As Double
  Do
    GetRand = 1 - Rnd - Rnd
  Loop Until Abs(GetRand) >= 0.1
End Function

Public Sub ExportNN(File As String)
  Dim a As Variant, b As Variant, c As Variant
  Dim i As Long, j As Long, k As Long
  
  'Saves the NN to a file. I avoided taking a few shortcuts here, keeping future expansion in mind.
  Dim FileNumb As Integer
  
  On Error Resume Next
  If (FileLen(File) = FileLen(File)) Then Kill (File)
  FileNumb = FreeFile
  Open File For Binary As FileNumb
    Put FileNumb, , "BrewNN v1.0.2       " & vbCrLf '20 Char description followed by carrage return
    
    ' Layers
    '########
    Put FileNumb, , LayerCount                      'Number of layers
    For i = 1 To LayerCount
      Put FileNumb, , CLng(0)                       'Pointers to layers (Set to 0 until position determined later)
    Next
    Put FileNumb, , CLng(0)                         'Pointer to MemBanks (Set to 0 until position determined later)
    For i = 1 To LayerCount
      a = Seek(FileNumb)
      Seek FileNumb, 24 + (2 * i)
      Put FileNumb, , CLng(a)                       'Save Layer position
      Seek FileNumb, CLng(a)
      Put FileNumb, , Layers(i).NeuronCount         'Number of neurons in current layer
      For j = 1 To Layers(i).NeuronCount
        Put FileNumb, , CLng(0)                     'Pointers to neurons in current layer (Set to 0 until position determined later)
      Next
      Put FileNumb, , CLng(UBound(Layers(i).AdjacentLayers) + 1) 'Number of adjacent layers
      For j = 0 To UBound(Layers(i).AdjacentLayers)              'Process Adjacent Layer info
        Put FileNumb, , CLng(Layers(i).AdjacentLayers(j))
      Next
      For j = 1 To Layers(i).NeuronCount            'Process neurons in current layer
        b = Seek(FileNumb)
        Seek FileNumb, CLng(a + 2 + (2 * j))
        Put FileNumb, , CLng(b)                             'Save neuron position
        Seek FileNumb, CLng(b)
        Put FileNumb, , Layers(i).Neurons(j).DendriteCount  'Number of dendrites in current neuron
        For k = 1 To Layers(i).Neurons(j).DendriteCount
          Put FileNumb, , CLng(0)                           'Pointers to dendrites in current neuron (Set to 0 until position determined later)
        Next
        Put FileNumb, , Layers(i).Neurons(j).BaseValue
        Put FileNumb, , Layers(i).Neurons(j).Error
        Put FileNumb, , Layers(i).Neurons(j).LearningCoefficient
        For k = 1 To Layers(i).Neurons(j).DendriteCount     'Process dendrites in current neuron
          c = Seek(FileNumb)
          Seek FileNumb, CLng(b + 2 + (2 * k))
          Put FileNumb, , CLng(c)                                   'Save dendrite position
          Seek FileNumb, CLng(c)
          Put FileNumb, , Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
          Put FileNumb, , CLng(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0))
          Put FileNumb, , CLng(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1))
        Next
      Next
    Next
    
    'MemBanks
    '########
    Put FileNumb, , MemBankCount                    'Number of membanks
    For i = 1 To MemBankCount
      Put FileNumb, , CLng(0)                       'Pointers to membanks (Set to 0 until position determined later)
    Next
    Put FileNumb, , CLng(0)                         'Pointer to layers (for continuity of structure :)
    For i = 1 To MemBankCount
      a = Seek(FileNumb)
      Seek FileNumb, 24 + (2 * i)
      Put FileNumb, , CLng(a)                       'Save membank position
      Seek FileNumb, CLng(a)
      Put FileNumb, , MemBanks(i).NeuronCount       'Number of neurons in current membank
      For j = 1 To MemBanks(i).NeuronCount
        Put FileNumb, , CLng(0)                     'Pointers to neurons in current membank (Set to 0 until position determined later)
      Next
      'Put FileNumb, , MemBanks(i).SlotCount         'Number of slots
      'For j = 0 To UBound(MemBanks(i).Slots)        'Process slot info
      '  Put FileNumb, , CLng(MemBanks(i).Slots(j))
      'Next
      For j = 1 To MemBanks(i).NeuronCount          'Process neurons in current membank
        b = Seek(FileNumb)
        Seek FileNumb, CLng(a + 2 + (2 * j))
        Put FileNumb, , CLng(b)                             'Save neuron position
        Seek FileNumb, CLng(b)
        Put FileNumb, , Layers(i).Neurons(j).DendriteCount  'Number of dendrites in current neuron
        For k = 1 To MemBanks(i).Neurons(j).DendriteCount
          Put FileNumb, , CLng(0)                           'Pointers to dendrites in current neuron (Set to 0 until position determined later)
        Next
        Put FileNumb, , MemBanks(i).Neurons(j).BaseValue
        Put FileNumb, , MemBanks(i).Neurons(j).Error
        Put FileNumb, , MemBanks(i).Neurons(j).LearningCoefficient
        For k = 1 To MemBanks(i).Neurons(j).DendriteCount   'Process dendrites in current membank
          c = Seek(FileNumb)
          Seek FileNumb, CLng(b + 2 + (2 * k))
          Put FileNumb, , CLng(c)                                   'Save dendrite position
          Seek FileNumb, CLng(c)
          Put FileNumb, , MemBanks(i).Neurons(j).Dendrites(k).ConnectionStrength
          Put FileNumb, , CLng(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(0))
          Put FileNumb, , CLng(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1))
        Next
      Next
    Next
  Close FileNumb
End Sub
