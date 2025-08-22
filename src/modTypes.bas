Attribute VB_Name = "modTypes"
Option Explicit

' 单条来源订单（已换算 usage_t）
Public Type FGOrder
    Company As String
    Product As String
    OrderID As String
    StartDate As Date
    EndDate As Date
    PlanQty As Double  ' 原始 plan order qty
    UsageT As Double   ' 换算后的吨数
End Type

' 批（按每周六锚点生成）
Public Type Batch
    BatchNo As Long
    anchor As Date
    FirstStart As Date
    LastStart As Date
    EffCapT As Double
    AllocatedT As Double
    windowDays As Long
    horizonDays As Long
End Type

' 分配明细行
Public Type AllocationLine
    Company As String
    Product As String
    OrderID As String
    StartDate As Date
    EndDate As Date
    BatchNo As Long
    anchor As Date
    UsageT As Double
    horizonDays As Long
End Type

Public Type tBatch
    BatchNo     As Long
    anchor      As Date
    FirstStart  As Date   ' ????
    LastStart   As Date   ' ????
    EffCapT     As Double
    AllocatedT  As Double
    windowDays  As Long
    horizonDays As Long
End Type

'Attribute VB_Name = "modTypes"


' === ?????? ===


' === ??????(?? Allocation ?) ===
Public Type tAlloc
    Company     As String
    Product     As String
    OrderID     As String
    StartDate   As Date
    EndDate     As Date
    BatchNo     As Long
    anchor      As Date
    UsageT      As Double
    horizonDays As Long
End Type

