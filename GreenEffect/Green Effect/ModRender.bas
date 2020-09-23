Attribute VB_Name = "ModGreen"

Public Type WType
    Name As String
    Attack As Double
    max As Double
    current As Double
    Misschance As Double
    Price As Double
End Type

Public Type AType
    Name As String
    max As Double
    current As Double
    Price As Double
End Type

Public Type Badie
    Name As String
    current As Double
    Attack As Double
    Misschance As Double
    Health As Double
End Type
    
Public Type Key
    Name As String
    KeyVal As Double
End Type

Public Type Items
    Name As String
    Description As String
    Extra As String
End Type



Public Weapon(1 To 10) As WType
Public Armour As AType

Public Monster(1 To 20, 1 To 20) As Badie

Public Keys(1 To 10) As Key
Public KeyCount As Integer

Public Item() As Items
Public ItemCount As Integer

Public PlayerHealth    'The Players Health
Public PlayerWeapon    'The Uses left of weapon
Public Castras          'Money

Sub Startup_WAA()   'The Startup (W)eapons (A)nd (A)rmour
    'This is the default weapon
    Weapon(1).Name = "Punch"    'The Name of the weapon
    Weapon(1).Attack = 1        'The attack value
    Weapon(1).max = -1          'The max uses of this weapon, -1 is no maximum
    Weapon(1).Misschance = 3    'The chance of missing out of ten
    Weapon(1).current = -1         'The amounted used, -1 is unlimited
    Weapon(1).Price = -1        'The selling price, -1 is N\A
    Armour.Name = "Clothes"
    Armour.max = 5
    Armour.Price = 0
    Armour.current = 5
    PlayerWeapon = 1
End Sub

Sub BuyArmour(Index)    'Buy armour, index is type
    If Index = 1 Then 'Cloth armour
        Armour.Name = "Cloth Armour"
        Armour.max = 10
        Armour.current = 10
        Armour.Price = 8
        
        Castras = Castras - 10
    ElseIf Index = 2 Then   'Leather armour
        Armour.Name = "Leather Armour"
        Armour.max = 20
        Armour.current = 20
        Armour.Price = 20
        
        Castras = Castras - 25
    ElseIf Index = 3 Then 'chainmail
        Armour.Name = "Chainmail"
        Armour.max = 40
        Armour.current = 40
        Armour.Price = 80
        
        Castras = Castras - 100
    End If
    Call FrmGreenEffect.DoProgress     'Update the Bars, of weapons, armour...
    Call DoSell         'Show the shopkeepers comment
End Sub

Sub buyweapon(Index)    'Buy a weapon, Index is type
    If Index = 1 Then
        For Check = 1 To 10
            If Weapon(Check).Name = "" Then
                Exit For
            End If
        Next
        Weapon(Check).Attack = 2
        Weapon(Check).max = -1
        Weapon(Check).Misschance = 3
        Weapon(Check).Name = "Big Stick"
        Weapon(Check).Price = 8
        Weapon(Check).current = -1
        
        Castras = Castras - 10
    ElseIf Index = 2 Then
        For Check = 1 To 10
            If Weapon(Check).Name = "" Then 'Search for free space
                Exit For
            End If
        Next
        Weapon(Check).Attack = 3
        Weapon(Check).max = -1
        Weapon(Check).Misschance = 3
        Weapon(Check).Name = "Club"
        Weapon(Check).Price = 16
        Weapon(Check).current = -1
        
        Castras = Castras - 20
    ElseIf Index = 3 Then
        For Check = 1 To 10
            If Weapon(Check).Name = "" Then 'Search fore free space
                Exit For
            End If
        Next
        Weapon(Check).Attack = 4
        Weapon(Check).max = -1
        Weapon(Check).Misschance = 3
        Weapon(Check).Name = "Axe"
        Weapon(Check).Price = 32
        Weapon(Check).current = -1
        
        Castras = Castras - 40
    ElseIf Index = 4 Then
        For Check = 1 To 10
            If Weapon(Check).Name = "" Then 'search for free space
                Exit For
            End If
        Next
        Weapon(Check).Attack = 2
        Weapon(Check).max = -1
        Weapon(Check).Misschance = 1
        Weapon(Check).Name = "Pike"
        Weapon(Check).Price = 60
        Weapon(Check).current = -1
        
        Castras = Castras - 75
    ElseIf Index = 5 Then
        For Check = 1 To 10
            If Weapon(Check).Name = "" Then 'Search for free space
                Exit For
            End If
        Next
        Weapon(Check).Attack = 5
        Weapon(Check).max = -1
        Weapon(Check).Misschance = 1
        Weapon(Check).Name = "Sword"
        Weapon(Check).Price = 120
        Weapon(Check).current = -1
        
        Castras = Castras - 150
    ElseIf Index = 6 Then
        For Check = 1 To 10
            If Weapon(Check).Name = "" Then 'Search for free space
                Exit For
            End If
        Next
        Weapon(Check).Attack = 2
        Weapon(Check).max = 10
        Weapon(Check).Misschance = 0
        Weapon(Check).Name = "Bow"
        Weapon(Check).Price = 400
        Weapon(Check).current = 0
        
        Castras = Castras - 500
    End If
    ask = MsgBox("Set new weapon to default?", vbYesNo)
    If ask = vbYes Then
        PlayerWeapon = Check
    End If
    Call FrmGreenEffect.DoProgress
    Call DoSell
End Sub

Sub WStatShow() 'This is for the weapon status form on its loadup
    FrmWStats.TVWeap.Nodes.Clear
    For OuterLoop = 1 To 10
        If Weapon(OuterLoop).Name <> "" Then
            Call FrmWStats.TVWeap.Nodes.Add(, , "Weapon" & Str(OuterLoop), "Name: " + Weapon(OuterLoop).Name)
            Call FrmWStats.TVWeap.Nodes.Add("Weapon" & Str(OuterLoop), tvwChild, "Attack" & Str(OuterLoop), "Attack:" + Str$(Weapon(OuterLoop).Attack))
            Call FrmWStats.TVWeap.Nodes.Add("Weapon" & Str(OuterLoop), tvwChild, "Max uses" & Str(OuterLoop), "Max uses:" + Str$(Weapon(OuterLoop).max))
            Call FrmWStats.TVWeap.Nodes.Add("Weapon" & Str(OuterLoop), tvwChild, "Missing(1-10)" & Str(OuterLoop), "Missing(1-10):" + Str$(Weapon(OuterLoop).Misschance))
            Call FrmWStats.TVWeap.Nodes.Add("Weapon" & Str(OuterLoop), tvwChild, "Used" & Str(OuterLoop), "Used:" + Str$(Weapon(OuterLoop).current))
        End If
    Next
    For OuterLoop = 1 To KeyCount
        Call FrmWStats.TVKey.Nodes.Add(, , "Key" & Str(OuterLoop), Keys(OuterLoop).Name)
        Call FrmWStats.TVKey.Nodes.Add("Key" & Str(OuterLoop), tvwChild, "KeyV" & Str(OuterLoop), Keys(OuterLoop).KeyVal)
    Next
End Sub

Sub SellWAA()   'This is for the sell weapon Form on its startup
    For OuterLoop = 1 To 10
        If Weapon(OuterLoop).Name = "" Then
            FrmShopSell.lblname(OuterLoop - 1).Caption = "Name: N\A"
            FrmShopSell.lblprice(OuterLoop - 1).Caption = "Price: N\A"
            FrmShopSell.cmdsell(OuterLoop - 1).Enabled = False
        ElseIf Weapon(OuterLoop).Name = "Punch" Then
            FrmShopSell.lblname(OuterLoop - 1).Caption = "Name: Punch"
            FrmShopSell.lblprice(OuterLoop - 1).Caption = "Price: N\A"
            FrmShopSell.cmdsell(OuterLoop - 1).Enabled = False
        Else
            FrmShopSell.lblname(OuterLoop - 1).Caption = "Name:" + Weapon(OuterLoop).Name
            If Weapon(OuterLoop).max = -1 Then
                FrmShopSell.lblprice(OuterLoop - 1) = "Price: " + Str$(Weapon(OuterLoop).Price)
            Else
                FrmShopSell.lblprice(OuterLoop - 1) = "Price:" + Str$(Weapon(OuterLoop).Price / Weapon(OuterLoop).max * (Weapon(OuterLoop).max - Weapon(OuterLoop).current))
            End If
            FrmShopSell.cmdsell(OuterLoop - 1).Enabled = True
        End If
    Next
    If Armour.Name = "" Then
        FrmShopSell.LblNameA = "Name: N\A"
        FrmShopSell.lblPriceA = "Price: N\A"
        FrmShopSell.CmdSellA.Enabled = False
    Else
        FrmShopSell.LblNameA = "Name: " + Armour.Name
        FrmShopSell.lblPriceA = "Price: " + Str$(Armour.Price / Armour.max * (Armour.max - Armour.current))
        FrmShopSell.CmdSellA.Enabled = True
    End If
End Sub

Sub SellW(ByVal Index As Integer)    'This is the sell weapon part
    If Index <> 1 Then
        If Weapon(Index).max = -1 Then
            Castras = Castras + Weapon(Index).Price
        Else
            Castras = Castras + (Weapon(Index).Price / Weapon(Index).max * (Weapon(Index).max - Weapon(Index).current))
        End If
        Weapon(Index).Attack = 0
        Weapon(Index).max = 0
        Weapon(Index).Misschance = 0
        Weapon(Index).Name = ""
        Weapon(Index).Price = 0
        Weapon(Index).current = 0
        PlayerWeapon = 1    'Stop you from using a weapon you may not have
    End If
    Call SellWAA    'Update the list
    Call FrmGreenEffect.DoProgress 'Update the Display
End Sub

Sub SellA() 'This is the sell armour part
    Castras = Castras + (Armour.Price / Armour.max * (Armour.max - Armour.current))
    Armour.max = 0
    Armour.Name = ""
    Armour.Price = 0
    Armour.current = 0
    
    Call SellWAA    'Update the List
    Call FrmGreenEffect.DoProgress 'Update the Display
End Sub

Sub AttackPTM(ByVal X As Integer, ByVal Y As Integer)    'Attack Subroutine (Person To Monster)
    If Weapon(PlayerWeapon).current > 0 Or Weapon(PlayerWeapon).max = -1 Then
        If Weapon(PlayerWeapon).current > 0 Then
            Weapon(PlayerWeapon).current = Weapon(PlayerWeapon).current - 1
        End If
        Prob = Rnd * 10
        If Prob > Weapon(PlayerWeapon).Misschance Then
            If Monster(X, Y).current - Weapon(PlayerWeapon).Attack < 0 Then
                HealthTake = Abs(Monster(X, Y).current - Weapon(PlayerWeapon).Attack)
                Monster(X, Y).current = 0
                Monster(X, Y).Health = Monster(X, Y).Health - HealthTake
            Else
                Monster(X, Y).current = Monster(X, Y).current - Weapon(PlayerWeapon).Attack
            End If
        End If
        If Monster(X, Y).Health <= 0 Then
            Call FrmGreenEffect.DrawDead
            Call ClearMonster(X, Y)
            Call FrmGreenEffect.ClearTrack
            Call FrmGreenEffect.GetMoney(Rnd * 100)
        End If
    End If
End Sub

Sub AttackMTP(ByVal X As Integer, ByVal Y As Integer)    'Attack Subroutine (Monster To Person)
    Prob = Rnd * 10
    If Prob > Monster(X, Y).Misschance Then
        If Armour.current - Monster(X, Y).Attack <= 0 Then
            HealthTake = Abs(Armour.current - Monster(X, Y).Attack)
            PlayerHealth = PlayerHealth - HealthTake
            Armour.current = 0
        Else
            Armour.current = Armour.current - Monster(X, Y).Attack
        End If
        If PlayerHealth <= 0 Then
            Call FrmGreenEffect.DoDie
        End If
    End If
End Sub

Sub ClearMonster(ByVal X As Integer, ByVal Y As Integer)
    Monster(X, Y).Attack = 0
    Monster(X, Y).Misschance = 0
    Monster(X, Y).Name = ""
    Monster(X, Y).current = 0
    Monster(X, Y).Health = 0
End Sub

Sub TempMonsters()
    For X = 1 To 20
        For Y = 1 To 20
            Monster(X, Y).Name = "Bob"
            Monster(X, Y).Attack = Rnd * 5
            Monster(X, Y).Misschance = 5 / 20 * (Rnd * X)
            Monster(X, Y).current = 10 / 20 * (Rnd * (X + Y) / 2)
            Monster(X, Y).Health = 100 / 20 * (Rnd * (X + Y) / 2)
        Next
    Next
End Sub
