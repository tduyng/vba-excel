Option Explicit
'Declare a new instance of our form changer class
'Khai báo bi?n s? du?c dùng cho class module CFormChanger
Dim mclsFormChanger As CFormChanger


Private Sub UserForm_Activate()
    'Giành m?t vùng nh? cho bi?n
    Set mclsFormChanger = New CFormChanger

    'Initialise to be like a 'standard' userform
    'Thi?t l?p các checkbox
    cbModal.Value = True
    cbCaption.Value = True
    cbCloseBtn.Value = True
    cbTaskBar.Value = True
    cbIcon.Value = False
    cbMaximize.Value = False
    cbMinimize.Value = False
    cbSizeable.Value = False
    cbSysmenu.Value = True
    cbTaskBar.Value = False
    cbSmallCaption.Value = False

    'Set the form changer to change this userform
    'Thi?t l?p bi?n cho UserForm
    Set mclsFormChanger.Form = Me

    'Make sure everything is in the right place to start with
    'Ch?c ch?n các control ? dúng v? trí, xin các b?n xem th? t?c UserForm_Resize
    UserForm_Resize

End Sub