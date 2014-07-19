Attribute VB_Name = "Module11"
Sub ANSI_to_Unicode()
'Created by Gevorg A. Galstyan

    Dim strFind As Variant
    Dim strRep As Variant
    strFind = Array(171, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 167, 166, 175, 170, 177, 176, 163, 58, 172, 173, 169, 168)
    strRep = Array(44, 1329, 1377, 1330, 1378, 1331, 1379, 1332, 1380, 1333, 1381, 1334, 1382, 1335, 1383, 1336, 1384, 1337, 1385, 1338, 1386, 1339, 1387, 1340, 1388, 1341, 1389, 1342, 1390, 1343, 1391, 1344, 1392, 1345, 1393, 1346, 1394, 1347, 1395, 1348, 1396, 1349, 1397, 1350, 1398, 1351, 1399, 1352, 1400, 1353, 1401, 1354, 1402, 1355, 1403, 1356, 1404, 1357, 1405, 1358, 1406, 1359, 1407, 1360, 1408, 1361, 1409, 1362, 1410, 1363, 1411, 1364, 1412, 1365, 1413, 1366, 1414, 171, 187, 1372, 1373, 1374, 1371, 1417, 1417, 45, 1418, 46, 1415)
    For i = 0 To 88
    
    Cells.Replace What:=ChrW(strFind(i)), Replacement:=ChrW(strRep(i)), LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    Next i
    
    Cells.Select
    Selection.Font.Name = "Sylfaen"
End Sub

Sub ConvertAll()
'Created by Gevorg A. Galstyan
Dim intSheets As Integer
intSheets = Sheets.Count
For i = 1 To intSheets
Sheets(i).Select
Application.Run "ANSI_to_Unicode"
Next i
Sheets(1).Select
MsgBox "Done", vbOKOnly, "Convertion Finished."
End Sub
