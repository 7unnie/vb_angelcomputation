Public Class Form1
    Dim QZ1, QZ2, QZ3, ATT, REC1, REC2, ACT1, ACT2, ACT3, EXAM, Prelim As Double



    Dim TQZ, TATT, TREC, TACT, TEXAM As Double

    Dim QUZ1, QUZ2, QUZ3, ATT1, RECI1, RECI2, ACTI1, ACTI2, ACTI3, EXAM1, Midterm As Double

    Dim TQZ1, TATT1, TREC1, TACT1, TEXAM1 As Double

    Dim QUIZ1, QUIZ2, QUIZ3, ATT2, RECIT1, RECIT2, ACTIV1, ACTIV2, ACTIV3, EXAM2, Finals As Double

    Dim TQZ2, TATT2, TREC2, TACT2, TEXAM2 As Double

    Dim TG As Double


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        QZ1 = Val(TextBox1.Text)
        QZ2 = Val(TextBox2.Text)
        QZ3 = Val(TextBox3.Text)
        ATT = Val(TextBox4.Text)
        REC1 = Val(TextBox5.Text)
        REC2 = Val(TextBox6.Text)
        ACT1 = Val(TextBox7.Text)
        ACT2 = Val(TextBox8.Text)
        ACT3 = Val(TextBox9.Text)
        EXAM = Val(TextBox10.Text)

        TQZ = ((QZ1 + QZ2 + QZ3) / 3) * 0.25
        TATT = ATT * 0.1
        TREC = ((REC1 + REC2) / 2) * 0.1
        TACT = ((ACT1 + ACT2 + ACT3) / 3) * 0.25
        TEXAM = EXAM * 0.3


        Prelim = (TQZ + TATT + TREC + TACT + TEXAM) * 0.3


        TextBox31.Text = Prelim


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        QUZ1 = Val(TextBox11.Text)
        QUZ2 = Val(TextBox12.Text)
        QUZ3 = Val(TextBox13.Text)
        ATT1 = Val(TextBox14.Text)
        RECI1 = Val(TextBox15.Text)
        RECI2 = Val(TextBox16.Text)
        ACTI1 = Val(TextBox17.Text)
        ACTI2 = Val(TextBox18.Text)
        ACTI3 = Val(TextBox19.Text)
        EXAM1 = Val(TextBox20.Text)

        TQZ1 = ((QUZ1 + QUZ2 + QUZ3) / 3) * 0.25
        TATT1 = ATT1 * 0.1
        TREC1 = ((RECI1 + REC2) / 2) * 0.1
        TACT1 = ((ACTI1 + ACT2 + ACT3) / 3) * 0.25
        TEXAM1 = EXAM1 * 0.3


        Midterm = (TQZ1 + TATT1 + TREC1 + TACT1 + TEXAM1) * 0.3


        TextBox32.Text = Midterm

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        QUIZ1 = Val(TextBox21.Text)
        QUIZ2 = Val(TextBox22.Text)
        QUIZ3 = Val(TextBox23.Text)
        ATT2 = Val(TextBox24.Text)
        RECIT1 = Val(TextBox25.Text)
        RECIT2 = Val(TextBox26.Text)
        ACTIV1 = Val(TextBox27.Text)
        ACTIV2 = Val(TextBox28.Text)
        ACTIV3 = Val(TextBox9.Text)
        EXAM2 = Val(TextBox30.Text)

        TQZ2 = ((QUIZ1 + QUIZ2 + QUIZ3) / 3) * 0.25
        TATT2 = ATT2 * 0.1
        TREC2 = ((RECIT1 + RECIT2) / 2) * 0.1
        TACT2 = ((ACTIV1 + ACTIV2 + ACTIV3) / 3) * 0.25
        TEXAM2 = EXAM2 * 0.3


        Finals = (TQZ2 + TATT2 + TREC2 + TACT2 + TEXAM2) * 0.4



        TextBox33.Text = Finals



    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        TG = (Prelim + Midterm + Finals)

        TextBox34.Text = TG

    End Sub


End Class
