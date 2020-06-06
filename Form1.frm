VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menambah Elemen Array"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
  Dim DynArray() 'Deklarasi array dinamik bernama
                 'DynArray
  ReDim DynArray(3) 'Deklarasi ulang sebanyak 4 elemen
  
  'Isi setiap elemen array, dan tampilkan
  For i = 0 To 2
    DynArray(i) = i
    MsgBox DynArray(i)
  Next i
  
  'Tambahkan 1 elemen ke DynArray
  ReDim Preserve DynArray(UBound(DynArray) + 1)
  
  'Isi elemen indeks ke-3 dengan 3
  DynArray(3) = 3
  
  'Tampilkan seluruh elemen setelah ditambah
  For i = 0 To 3
    MsgBox DynArray(i)
  Next i
End Sub

'Hanya batas teratas dari dimensi terakhir dalam sebuah 'array multidimensional yang dapat diganti ketika Anda 'menggunakan kata kunci Preserve keyword ini.

'Jika Anda mengubah elemen dimensi lainnya, atau indeks 'yang lebih rendah dari dimensi yang terakhir, sebuah 'error saat run-time akan terjadi.
'Jadi, Anda dapat menggunakan coding seperti ini:

Private Sub Command2_Click()
Dim i As Integer, j As Integer
  Dim Matrix() As String 'Deklarasikan array dinamik
                         'Matrix
  ReDim Matrix(2, 3) As String
  'Deklarasi Matrix sebagai array Multidimensi
  'Isi elemen array Matrix...
  Matrix(1, 1) = "Baris satu kolom satu"
  Matrix(1, 2) = "Baris satu kolom dua"
  Matrix(1, 3) = "Baris satu kolom tiga"
  Matrix(2, 1) = "Baris dua kolom satu"
  Matrix(2, 2) = "Baris dua kolom dua"
  Matrix(2, 3) = "Baris dua kolom tiga"
    
  'Tampilkan semua elemen array Matrix
  For i = 1 To 2
    For j = 1 To 3  'Mula-mula masih 3...
      MsgBox Matrix(i, j)
    Next j
  Next i
   
  'Tambahkan satu elemen di dimensi yang terakhir
  '(kanan)
  ReDim Preserve Matrix(2, UBound(Matrix, 2) + 1)
  'Tapi, Anda tidak dapat menggunakan coding berikut:
  'ReDim Preserve Matrix(UBound(Matrix, 1) + 1, 10)
  'karena akan menyebabkan error pada saat run-time
  
  'Sekarang, array ini menjadi Matrix(2, 4)
  'Isi nilai elemen yang ditambahkan ini
  Matrix(1, 4) = "Baris satu kolom empat (baru)"
  Matrix(2, 4) = "Baris dua kolom empat (baru)"

  'Tampilkan semua elemen array Matrix
  For i = 1 To 2
    For j = 1 To 4  'Sekarang sudah menjadi 4...
      MsgBox Matrix(i, j)
    Next j
  Next i
End Sub

