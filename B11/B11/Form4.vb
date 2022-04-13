Public Class Form4
    Dim strg1, strg2 As String
    Dim jmlh, jmlh1, strg As Integer
    Dim ip, ip1 As TimeSpan
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Form2.Close()
        'TabelTZ.Close()
        strg1 = "True"
        strg2 = "False"
        ip = DTend.Value - DTstart.Value
        Form5.DTadd1.Value = DateAdd("d", -1, DTstart.Value)
        Form5.DTadd2.Value = DateAdd("d", 1, DTend.Value)
        Form5.Show()
        Select Case ComboBox1.Text

            Case "GEPOST" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg2)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)



                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()
                        ''Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                        'End If
                    Else Label.Text = "Salah memasukkan tanggal. 
                                       Perhatikan tanggal mulai dan selesainya"
                    End If
                End If


            Case "PPIC" : ip = DTend.Value - DTstart.Value

                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg2)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()
                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "LAMINAR2" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg2)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)

                If Val(ip.Days) >= 0 Then

                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()
                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "LAMINAR1" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg2)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)

                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()
                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "Emulsi2" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg2)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "EMULSI1" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg2)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "EDSPOST" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg2)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "EDSINO" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg2)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "EDSPANEN" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg2)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg2)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If
            Case "GEINO" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg2)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "KORIDOR" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg2)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "GEPANEN" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg2)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                  Perhatikan tanggal mulai dan selesainya"
                End If

            Case "PASSAGE" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg2)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)


                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()
                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                  Perhatikan tanggal mulai dan selesainya"
                End If

            Case "PRA" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)



                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg2)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal.
                                  Perhatikan tanggal mulai dan selesainya"
                End If

            Case "ALAT" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg2)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                  Perhatikan tanggal mulai dan selesainya"
                End If

            Case "STERILISASI" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg2)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "PERSIAPAN" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg2)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "INO2" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg2)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "INO1" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg2)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "POST2" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg2)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                  Perhatikan tanggal mulai dan selesainya"
                End If
            Case "POST1" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg2)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                  Perhatikan tanggal mulai dan selesainya"
                End If

            Case "PANEN1" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg2)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                Perhatikan tanggal mulai dan selesainya"
                End If

            Case "PANEN2" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg2)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "INAKTIF1" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)


                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg2)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                    Perhatikan tanggal mulai dan selesainya"
                End If

            Case "INAKTIF2" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg2)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)

                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                    Perhatikan tanggal mulai dan selesainya"
                End If


            Case "KORIDORFADV" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg2)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then

                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "CELLFADV" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg2)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then

                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "PANENFADV" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg2)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then

                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "INOFADV" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg2)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg1)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then

                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "POSTFADV" : ip = DTend.Value - DTstart.Value
                Form2.CrystalTabel1.SetParameterValue("awal", DTstart.Text)
                Form2.CrystalTabel1.SetParameterValue("Akhir", Form5.DTadd2.Value)

                Form2.CrystalTabel1.SetParameterValue("GEPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("PPIC", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR2", strg1)
                Form2.CrystalTabel1.SetParameterValue("LAMINAR1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI2", strg1)
                Form2.CrystalTabel1.SetParameterValue("EMULSI1", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPOST", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("EDSPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEINO", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDOR", strg1)
                Form2.CrystalTabel1.SetParameterValue("GEPANEN", strg1)
                Form2.CrystalTabel1.SetParameterValue("PASSAGE", strg1)
                Form2.CrystalTabel1.SetParameterValue("PRA", strg1)
                Form2.CrystalTabel1.SetParameterValue("ALAT", strg1)
                Form2.CrystalTabel1.SetParameterValue("STERILISASI", strg1)
                Form2.CrystalTabel1.SetParameterValue("PERSIAPAN", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INO1", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST2", strg1)
                Form2.CrystalTabel1.SetParameterValue("POST1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN1", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANEN2", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF1", strg1)
                Form2.CrystalTabel1.SetParameterValue("INAKTIF2", strg1)
                Form2.CrystalTabel1.SetParameterValue("KORIDORFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("CELFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("PANENFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("INOFADV", strg1)
                Form2.CrystalTabel1.SetParameterValue("POSTFADV", strg2)
                'TextBox1.Text = bulankeangka(Mid(DTend.Text, 4, jmlh1 - 8)) - bulankeangka(Mid(DTstart.Text, 4, jmlh - 8))

                If Val(ip.Days) >= 0 Then

                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        Form2.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/032"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                  Perhatikan tanggal mulai dan selesainya"
                End If


            Case "BIO/COLD/037"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                 Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/038"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/039"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                   Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/040"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/041"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/042"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/043"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg1)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg2)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                Perhatikan tanggal mulai dan selesainya"
                End If

            Case "BIO/COLD/044"
                TabelTZ.CrystalReport21.SetParameterValue("start", DTstart.Text)
                TabelTZ.CrystalReport21.SetParameterValue("end", Form5.DTadd2.Text)
                TabelTZ.CrystalReport21.SetParameterValue("032", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("037", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("038", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("039", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("040", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("041", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("042", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("043", strg2)
                TabelTZ.CrystalReport21.SetParameterValue("044", strg1)

                If Val(ip.Days) >= 0 Then
                    If Val(ip.Days) <= 30 Then
                        Label.Text = ""
                        TabelTZ.Show()

                        'Else Label.Text = "Data yang ditampilkan :" + ip.Days + ", Maksimal 30 hari"
                    End If
                Else Label.Text = "Salah memasukkan tanggal. 
                                    Perhatikan tanggal mulai dan selesainya"
                End If

        End Select
        Form5.Close()
        If ComboBox1.Text = "Ruangan" Then Label.Text = "Pilih ruangan terlebih dahulu !!!"
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label10.Text = "*Data tabel yang dapat ditampilkan 
                        maksimal 30 hari"

        If Val(Label9.Text) Mod 2 = 0 Then
            ComboBox1.Items.Add("GEPOST")
            ComboBox1.Items.Add("PPIC")
            ComboBox1.Items.Add("LAMINAR2")
            ComboBox1.Items.Add("LAMINAR1")
            ComboBox1.Items.Add("EMULSI2")
            ComboBox1.Items.Add("EMULSI1")
            ComboBox1.Items.Add("EDSPOST")
            ComboBox1.Items.Add("EDSINO")
            ComboBox1.Items.Add("EDSPANEN")
            ComboBox1.Items.Add("GEINO")
            ComboBox1.Items.Add("KORIDOR")
            ComboBox1.Items.Add("GEPANEN")
            ComboBox1.Items.Add("PASSAGE")
            ComboBox1.Items.Add("PRA")
            ComboBox1.Items.Add("ALAT")
            ComboBox1.Items.Add("STERILISASI")
            ComboBox1.Items.Add("PERSIAPAN")
            ComboBox1.Items.Add("INO2")
            ComboBox1.Items.Add("INO1")
            ComboBox1.Items.Add("POST2")
            ComboBox1.Items.Add("POST1")
            ComboBox1.Items.Add("PANEN2")
            ComboBox1.Items.Add("PANEN1")
            ComboBox1.Items.Add("INAKTIF1")
            ComboBox1.Items.Add("INAKTIF2")
            ComboBox1.Items.Add("KORIDORFADV")
            ComboBox1.Items.Add("CELLFADV")
            ComboBox1.Items.Add("PANENFADV")
            ComboBox1.Items.Add("INOFADV")
            ComboBox1.Items.Add("POSTFADV")
            ComboBox1.Items.Add("WTCL-012")
            ComboBox1.Items.Add("WTCL-024")

        End If
        If Val(Label9.Text) Mod 2 = 1 Then
            ComboBox1.Items.Add("BIO/COLD/032")
            ComboBox1.Items.Add("BIO/COLD/037")
            ComboBox1.Items.Add("BIO/COLD/038")
            ComboBox1.Items.Add("BIO/COLD/039")
            ComboBox1.Items.Add("BIO/COLD/040")
            ComboBox1.Items.Add("BIO/COLD/041")
            ComboBox1.Items.Add("BIO/COLD/042")
            ComboBox1.Items.Add("BIO/COLD/043")
            ComboBox1.Items.Add("BIO/COLD/044")
            ComboBox1.Items.Add("WTCL012")
            ComboBox1.Items.Add("WTCL024")
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        grafik.Close()
        GrafikTZ.Close()
        Form5.Show()
        Form5.DTgraf2.Value = DateAdd("d", 3, Me.DTgraf1.Value)
        strg1 = "True"
        strg2 = "False"
        Select Case ComboBox1.Text
            Case "GEPOST"
                grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg2)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)

                grafik.Show()
            Case "PPIC" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg2)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "LAMINAR2" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg2)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "LAMINAR1" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg2)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "EMULSI2" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg2)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "EMULSI1" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg2)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "EDSPOST" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg2)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "EDSINO" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg2)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "EDSPANEN" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg2)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "GEINO" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg2)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "KORIDOR" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg2)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "GEPANEN" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg2)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "PASSAGE" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg2)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "PRA"
                grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg2)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "ALAT" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg2)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "STERILISASI" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg2)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "PERSIAPAN" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg2)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "INO2" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg2)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "INO1" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg2)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "POST2" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg2)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "POST1" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg2)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "PANEN1" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg2)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "PANEN2" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg2)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "INAKTIF1" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg2)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()
            Case "INAKTIF2" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg2)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "KORIDORFADV" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg2)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "CELLFADV" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg2)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "PANENFADV" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg2)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "INOFADV" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg2)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "POSTFADV" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg2)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "WTCL012" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg2)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg1)
                grafik.Show()

            Case "WTCL024" : grafik.CrystalGrafik1.SetParameterValue("start", DTgraf1.Text)
                grafik.CrystalGrafik1.SetParameterValue("end", Form5.DTgraf2.Text)

                grafik.CrystalGrafik1.SetParameterValue("GEPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PPIC", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("LAMINAR1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EMULSI1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPOST", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("EDSPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEINO", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDOR", strg1)
                grafik.CrystalGrafik1.SetParameterValue("GEPANEN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PASSAGE", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PRA", strg1)
                grafik.CrystalGrafik1.SetParameterValue("ALAT", strg1)
                grafik.CrystalGrafik1.SetParameterValue("STERILISASI", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PERSIAPAN", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INO1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POST1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANEN2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF1", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INAKTIF2", strg1)
                grafik.CrystalGrafik1.SetParameterValue("KORIDORFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("CELLFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("PANENFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("INOFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("POSTFADV", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL012", strg1)
                grafik.CrystalGrafik1.SetParameterValue("WTCL024", strg2)
                grafik.Show()



            Case "BIO/COLD/032" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/037" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/038" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/039" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/040" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/041" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/042" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/043" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg2)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
                GrafikTZ.Show()
            Case "BIO/COLD/044" : GrafikTZ.CrystalReport31.SetParameterValue("start", DTgraf1.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.DTgraf2.Text)
                GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
                GrafikTZ.CrystalReport31.SetParameterValue("044", strg2)
                GrafikTZ.Show()
        End Select
        Form5.Close()
        If ComboBox1.Text = "Ruangan" Then Label.Text = "Pilih ruangan terlebih dahulu !!!"
    End Sub

    'Function jumlahhari(ByVal c As String)
    '    Dim output As Integer
    '    Select Case Mid(c, 4, Len(c) - 8)
    '        Case "January" : output = 31
    '        Case "February" : If Mid(c, Len(c) - 3) Mod 4 = 0 Then
    '                output = 29
    '            Else output = 28
    '            End If
    '        Case "March" : output = 31
    '        Case "April" : output = 31
    '        Case "May" : output = 31
    '        Case "June" : output = 31
    '        Case "July" : output = 31
    '        Case "August" : output = 31
    '        Case "September" : output = 31
    '        Case "October" : output = 31
    '        Case "" : output = 31
    '        Case "January" : output = 31

    '    End Select
    'End Function
    'Function bulankeangka(ByVal c As String)
    '    Dim output As Integer
    '    Select Case c
    '        Case "January" : output = 1
    '        Case "February" : output = 2
    '        Case "March" : output = 3
    '        Case "April" : output = 4
    '        Case "May" : output = 5
    '        Case "June" : output = 6
    '        Case "July" : output = 7
    '        Case "August" : output = 8
    '        Case "September" : output = 9
    '        Case "October" : output = 10
    '        Case "November" : output = 11
    '        Case "December" : output = 12
    '    End Select
    '    Return output
    'End Function
    'Function angkakebulan(ByVal c As Integer)
    '    Dim output1 As String
    '    Select Case c
    '        Case 1 : output1 = "January"
    '        Case 2 : output1 = "February"
    '        Case 3 : output1 = "March"
    '        Case 4 : output1 = "April"
    '        Case 5 : output1 = "May"
    '        Case 6 : output1 = "June"
    '        Case 7 : output1 = "July"
    '        Case 8 : output1 = "August"
    '        Case 9 : output1 = "September"
    '        Case 10 : output1 = "October"
    '        Case 11 : output1 = "November"
    '        Case 12 : output1 = "December"

    '    End Select
    '    Return output1
    'End Function
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class