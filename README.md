# VBA-challenge

Here is my submission for the VBA Challenge on Stock Data which includes the following: 

1. The Excel file used to run the code.
2. A file with the VBA script with all the lines of code I used to complete the assingment.
3. Three (3) screenshots of the sheets - first page only


Please note:
I got assistance from the Learning Assistants on Slack (AskBCS) and they gave me guidance on lines 102 to 104 to find the maximum and minimum values in the column to calculate values for Greatest % increase and Greatest % decrease. Here are the lines:
    
    ws.Range("R2").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastsumrow))
    ws.Range("R3").Value = WorksheetFunction.Min(ws.Range("L2:L" & lastsumrow))
    ws.Range("R4").Value = WorksheetFunction.Max(ws.Range("M2:M" & lastsumrow))

Thank you!

