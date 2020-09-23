Attribute VB_Name = "ModText"
Sub NoSell()
    FrmGreenEffect.LblMessage.Caption = "Talking to Shop Keeper"
    FrmGreenEffect.LblText.Caption = "Sorry you can't afford that, why don't you come back later on..."
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoSell()
    FrmGreenEffect.LblMessage.Caption = "Talking to Shop Keeper"
    FrmGreenEffect.LblText.Caption = "Thank You... Would you be interested in anything else?"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoBCase()
    FrmGreenEffect.LblMessage.Caption = "Looking at a BookCase"
    FrmGreenEffect.LblText.Caption = "Such an interesting Bookcase, shame i can't read the writing"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoCase()
    FrmGreenEffect.LblMessage.Caption = "Looking at a chest"
    FrmGreenEffect.LblText.Caption = "Lots of different objects, wish i had some"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoBed()
    FrmGreenEffect.LblMessage.Caption = "Looking at the Bed"
    FrmGreenEffect.LblText.Caption = "Such a comfortable Bed, but i don't have time to rest"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoFplace()
    FrmGreenEffect.LblMessage.Caption = "Looking at a FirePlace"
    FrmGreenEffect.LblText.Caption = "This is very warm, can't go to near, might burn myself"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

