Attribute VB_Name = "Module11"
Public Sub Func()
    Dim doc As AssemblyDocument
    Set doc = ThisApplication.ActiveDocument
    
    Dim holeList As New Collection
    Dim normalList As New Collection
    Dim pointList As New Collection
    
    For Each occ In doc.ComponentDefinition.Occurrences
        If InStr(occ.Name, "hole") <> 0 Then
            Dim hole As ComponentOccurrence
            Set hole = occ
            holeList.Add hole, hole.Name
        End If
    Next
    
    For i = 0 To holeList.Count Step 1
        For Each currentHole In holeList
        If i = getNumber(currentHole.Name) Then
            For Each oFace In currentHole.SurfaceBodies(1).Faces
                If oFace.SurfaceType = kPlaneSurface Then
                    If oFace.IsParamReversed = False Then
                        Dim oNormal As UnitVector
                        Set oNormal = oFace.Geometry.normal
                        normalList.Add oNormal, currentHole.Name
                        
                        Dim oPoint As point
                        Set oPoint = oFace.Geometry.RootPoint
                        pointList.Add oPoint, currentHole.Name
                    End If
                End If
            Next
        End If
        Next
    Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile("vba.json", True, False)

    Fileout.WriteLine "{"
    Fileout.WriteLine "    " & Chr(34) & "acc_bias" & Chr(34) & ": ["
    Fileout.WriteLine "        -0.1153,"
    Fileout.WriteLine "        0.06581,"
    Fileout.WriteLine "        -0.09107"
    Fileout.WriteLine "    ],"
    Fileout.WriteLine "    " & Chr(34) & "acc_scale" & Chr(34) & ": ["
    Fileout.WriteLine "        0.9995,"
    Fileout.WriteLine "        1,"
    Fileout.WriteLine "        0.9978"
    Fileout.WriteLine "    ],"
    Fileout.WriteLine "    " & Chr(34) & "device_class" & Chr(34) & ": " & Chr(34) & "controller"","
    Fileout.WriteLine "    " & Chr(34) & "device_pid" & Chr(34) & ": 8210,"
    Fileout.WriteLine "    " & Chr(34) & "device_serial_number" & Chr(34) & ": " & Chr(34) & "LHR-F539B946" & Chr(34) & ","
    Fileout.WriteLine "    " & Chr(34) & "device_vid" & Chr(34) & ": 10462,"
    Fileout.WriteLine "    " & Chr(34) & "gyro_bias" & Chr(34) & ": ["
    Fileout.WriteLine "        -0.02063,"
    Fileout.WriteLine "        -0.0001659,"
    Fileout.WriteLine "        -0.03915"
    Fileout.WriteLine "    ],"
    Fileout.WriteLine "    " & Chr(34) & "gyro_scale" & Chr(34) & ": ["
    Fileout.WriteLine "        1.0,"
    Fileout.WriteLine "        1.0,"
    Fileout.WriteLine "        1.0"
    Fileout.WriteLine "    ],"
    Fileout.WriteLine "    " & Chr(34) & "htcComposeTime" & Chr(34) & ": " & Chr(34) & "2016-05-05 03:24:53.673000" & Chr(34) & ","
    Fileout.WriteLine "    " & Chr(34) & "lighthouse_config" & Chr(34) & ": {"
    Fileout.WriteLine "        " & Chr(34) & "channelMap" & Chr(34) & ": ["
    
    For j = 0 To holeList.Count - 2 Step 1
        Fileout.WriteLine "            " & j & ","
    Next
    Fileout.WriteLine "            " & holeList.Count - 1
    Fileout.WriteLine "        ],"
    Fileout.WriteLine "        " & Chr(34) & "modelNormals" & Chr(34) & ": ["

        Fileout.WriteLine "            ["
        Fileout.WriteLine "                " & Replace(-normalList(1).X, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(-normalList(1).Y, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(-normalList(1).Z, ",", ".", 1)
        Fileout.Write "            ]"
    
    For i = 2 To 24 Step 1
        Fileout.WriteLine ","
        Fileout.WriteLine "            ["
        Fileout.WriteLine "                " & Replace(-normalList(i).X, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(-normalList(i).Y, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(-normalList(i).Z, ",", ".", 1)
        Fileout.Write "            ]"
    Next
    
    Fileout.WriteLine "            ],"
    Fileout.WriteLine "        " & Chr(34) & "modelPoints" & Chr(34) & ": ["
    
    
        Fileout.WriteLine "            ["
        Fileout.WriteLine "                " & Replace(pointList(1).X * 0.01, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(pointList(1).Y * 0.01, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(pointList(1).Z * 0.01, ",", ".", 1)
        Fileout.Write "            ]"
    
    For i = 2 To 24 Step 1
        Fileout.WriteLine ","
        Fileout.WriteLine "            ["
        Fileout.WriteLine "                " & Replace(pointList(i).X * 0.01, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(pointList(i).Y * 0.01, ",", ".", 1) & ", "
        Fileout.WriteLine "                " & Replace(pointList(i).Z * 0.01, ",", ".", 1)
        Fileout.Write "            ]"
    Next
        
        
    Fileout.WriteLine "]"
    Fileout.WriteLine "    },"
    Fileout.WriteLine "    " & Chr(34) & "manufacturer" & Chr(34) & ": " & Chr(34) & "HTC" & Chr(34) & ","
    Fileout.WriteLine "    " & Chr(34) & "mb_serial_number" & Chr(34) & ": " & Chr(34) & "42FM164P13809" & Chr(34) & ","
    Fileout.WriteLine "    " & Chr(34) & "model_number" & Chr(34) & ": " & Chr(34) & "Vive Controller MV" & Chr(34) & ","
    Fileout.WriteLine "    " & Chr(34) & "render_model" & Chr(34) & ": " & Chr(34) & "vr_controller_vive_1_5" & Chr(34) & ","
    Fileout.WriteLine "    " & Chr(34) & "revision" & Chr(34) & ": 1,"
    Fileout.WriteLine "    " & Chr(34) & "trackref_from_head" & Chr(34) & ": ["
    Fileout.WriteLine "        0,"
    Fileout.WriteLine "        0.7253744006156921,"
    Fileout.WriteLine "        -0.6883545517921448,"
    Fileout.WriteLine "        0,"
    Fileout.WriteLine "        0,"
    Fileout.WriteLine "        0.07100000232458115,"
    Fileout.WriteLine "        -0.03099999949336052"
    Fileout.WriteLine "    ],"
    Fileout.WriteLine "    " & Chr(34) & "trackref_from_imu" & Chr(34) & ": ["
    Fileout.WriteLine "        0,"
    Fileout.WriteLine "        0,"
    Fileout.WriteLine "        0,"
    Fileout.WriteLine "        1,"
    Fileout.WriteLine "        -0.006630099844187498,"
    Fileout.WriteLine "        -0.05046970024704933,"
    Fileout.WriteLine "        -0.023625019937753677"
    Fileout.WriteLine "    ],"
    Fileout.WriteLine "    " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Lighthouse_HMD" & Chr(34)
    Fileout.WriteLine "}"
   
    Fileout.Close
    MsgBox ("���������� ��������: " & normalList.Count & " ���������� �����: " & pointList.Count)
    
    
End Sub
Function getNumber(text As String) As Integer
    Dim oTestArray() As String
    oTestArray = Split(text, ":")
    getNumber = oTestArray(1)
End Function

