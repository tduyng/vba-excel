
Private Sub UserForm_Initialize()
Application.OnTime Now + TimeValue("00:00:03"), "LogoSTOP"
End Sub

Sub LogoSTOP()
LogoSplashScreen.Hide
End Sub
