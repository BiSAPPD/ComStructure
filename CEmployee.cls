VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pUid As String
Private pSpec As String
Private pSex As String
Private pDir As String
Private pReg As String
Private pRole As String
Private pChief As String
Private pNewOne As Boolean
Private pLost As Boolean
Private pAbsent As Boolean
Private pVacancy As Integer
Private pCA As String
Private pActive As Integer
Private pMail As String
Private pSubRole As String

Public Property Get SubRole() As String
    SubRole = pSubRole
End Property
Public Property Let SubRole(Value As String)
    pSubRole = Value
End Property
Public Property Get Mail() As String
    Mail = pMail
End Property
Public Property Let Mail(Value As String)
    pMail = Value
End Property
Public Property Get Active() As Integer
    Active = pActive
End Property
Public Property Let Active(Value As Integer)
    pActive = Value
End Property
Public Property Get CA() As String
    CA = pCA
End Property
Public Property Let CA(Value As String)
    pCA = Value
End Property
Public Property Get Vacancy() As Integer
    Vacancy = pVacancy
End Property
Public Property Let Vacancy(Value As Integer)
    pVacancy = Value
End Property
Public Property Get Uid() As String
    Uid = pUid
End Property
Public Property Let Uid(Value As String)
    pUid = Value
End Property
Public Property Get name() As String
    name = pName
End Property
Public Property Let name(Value As String)
    pName = Value
End Property
Public Property Get Spec() As String
    Spec = pSpec
End Property
Public Property Let Spec(Value As String)
    pSpec = Value
End Property
Public Property Get Sex() As String
    Sex = pSex
End Property
Public Property Let Sex(Value As String)
    pSex = Value
End Property
Public Property Get Dir() As String
    Dir = pDir
End Property
Public Property Let Dir(Value As String)
    pDir = Value
End Property
Public Property Get Role() As String
    Role = pRole
End Property
Public Property Let Role(Value As String)
    pRole = Value
End Property
Public Property Get Reg() As String
    Reg = pReg
End Property
Public Property Let Reg(Value As String)
    pReg = Value
End Property
Public Property Get Chief() As String
    Chief = pChief
End Property
Public Property Let Chief(Value As String)
    pChief = Value
End Property
Public Property Get NewOne() As Boolean
    NewOne = pNewOne
End Property
Public Property Let NewOne(Value As Boolean)
    pNewOne = Value
End Property
Public Property Get Lost() As Boolean
    Lost = pLost
End Property
Public Property Let Lost(Value As Boolean)
    pLost = Value
End Property
Public Property Get Absent() As Boolean
    Absent = pAbsent
End Property
Public Property Let Absent(Value As Boolean)
    pAbsent = Value
End Property
