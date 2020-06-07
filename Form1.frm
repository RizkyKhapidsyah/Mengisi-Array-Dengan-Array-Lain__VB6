VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim intX As Integer   'Deklarasi counter array
   
   'Deklarasi dan isi sebuah array integer.
   Dim countersA(5) As Integer
   For intX = 0 To 4
       countersA(intX) = intX
   Next intX
  
   'Deklarasi dan isi sebuah array string.
   Dim countersB(5) As String
   For intX = 0 To 4
       countersB(intX) = "Teks ke-" & intX
   Next intX
   
   Dim arrX(2) As Variant   'Deklarasi sebuah array.

   'Dalam hal ini kita tidak menggunakan indeks ke-0
   arrX(1) = countersA() 'Isi arrX(1) dengan array CountersA
   arrX(2) = countersB() 'Isi arrX(2) dengan array CountersB

   'Tampilkan sebuah elemen dari setiap arrX
   MsgBox arrX(1)(2)  'Menghasilkan 2
   MsgBox arrX(2)(3)  'Menghasilkan "Teks ke-3"
   
End Sub
