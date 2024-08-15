# VBA-Challenge 

The Process I Took to Create a VBA Script to Accomplish the Task

When tasked with analyzing stock data and identifying key metrics such as the greatest percentage increase, greatest percentage decrease, and the greatest total stock volume across multiple worksheets, I knew the most efficient way to handle this was by creating a VBA macro. Here’s the step-by-step process I followed to develop and implement the solution:

Step 1: Defining the Task and Objectives
Before I even opened Excel, I took a moment to understand the requirements. The goal was clear: I needed to analyze stock data across several worksheets (each representing a quarter) and identify three key metrics for each quarter:
-Greatest Percentage Increase
-Greatest Percentage Decrease
-Greatest Total Volume

Additionally, I had to implement conditional formatting to automatically highlight positive percentage changes in green and negative percentage changes in red. I also wanted the macro to be dynamic, meaning it would allow users to input a stock ticker and get relevant performance data on demand. The results had to be outputted clearly next to the data, just as shown in the provided image.

Step 2: Writing the Main VBA Script
I started by writing the main subroutine called `CalculateStockMetrics`. The core functionality of this macro was to loop through all worksheets in the workbook, calculate the key metrics, and then display the results for each quarter directly on the corresponding worksheet.

Looping Through Worksheets
The first task was to loop through each worksheet. This allowed the macro to handle all quarters automatically, without needing to run the script on each worksheet manually.

Calculating the Metrics
For each worksheet, I needed to loop through the rows of stock data. I focused on identifying:
- The stock with the greatest percentage increase.
- The stock with the greatest percentage decrease.
- The stock with the greatest total volume.

I used variables to store these values as I iterated through the data. I compared each stock's performance, updating the variables whenever a higher increase, larger decrease, or greater volume was found. Once the loop was completed, I had the tickers and values for the top-performing stocks ready.

Displaying the Results
Next, I ensured the results were outputted in the specified cells, just as shown in the provided example image. I chose cells in columns adjacent to the data to display the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume, along with the corresponding tickers and values.

Here’s the main macro code I wrote:

vba
Sub CalculateStockMetrics()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    Dim i As Long

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        greatestIncrease = -9999
        greatestDecrease = 9999
        greatestVolume = 0
        
        ' Loop through each row to find the greatest increase, decrease, and volume
        For i = 2 To lastRow
            ticker = ws.Cells(i, 7).Value

            ' Find greatest percentage increase
            If ws.Cells(i, 9).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(i, 9).Value
                increaseTicker = ticker
            End If
            
            ' Find greatest percentage decrease
            If ws.Cells(i, 9).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(i, 9).Value
                decreaseTicker = ticker
            End If
            
            ' Find greatest total volume
            If ws.Cells(i, 10).Value > greatestVolume Then
                greatestVolume = ws.Cells(i, 10).Value
                volumeTicker = ticker
            End If
        Next i
        
        ' Output results on the worksheet
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(2, 15).Value = increaseTicker
        ws.Cells(3, 15).Value = decreaseTicker
        ws.Cells(4, 15).Value = volumeTicker
        
        ws.Cells(2, 16).Value = greatestIncrease & "%"
        ws.Cells(3, 16).Value = greatestDecrease & "%"
        ws.Cells(4, 16).Value = greatestVolume
        
        ' Apply conditional formatting for positive and negative changes
        Call ApplyConditionalFormatting(ws, lastRow)
    Next ws

End Sub
```

Step 3: Adding Conditional Formatting**
After completing the calculations, I focused on applying conditional formatting. I wanted positive percentage changes to be highlighted in green and negative percentage changes in red. To achieve this, I wrote a subroutine called `ApplyConditionalFormatting`.

Here’s how the formatting works:
- I applied the formatting to the range of percentage change values dynamically based on the number of rows.
- I cleared any existing formatting before applying the new rules to ensure consistency.

Here’s the conditional formatting subroutine:

```vba
Sub ApplyConditionalFormatting(ws As Worksheet, lastRow As Long)

    Dim rng As Range
    Set rng = ws.Range("I2:I" & lastRow)

    ' Clear any existing formatting
    rng.FormatConditions.Delete

    ' Apply green fill for positive changes
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Interior.Color = RGB(0, 255, 0) ' Green
    End With

    ' Apply red fill for negative changes
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Interior.Color = RGB(255, 0, 0) ' Red
    End With

End Sub
```

Step 4: Testing and Validating the Macro
Once the code was written, I ran the macro by pressing `Alt + F8`, selecting `CalculateStockMetrics`, and clicking `Run`. I tested the macro on the dataset, making sure that it correctly identified the top-performing stocks and applied the proper formatting.

I checked each worksheet to ensure the results were accurate and that the conditional formatting was correctly applied. Everything worked as expected: the stocks with the greatest percentage increase, decrease, and volume were identified, and the percentage changes were visually distinguishable by their colors.

Step 6: Implementing User Input Capability
To make the macro even more dynamic, I planned to add a feature that allowed users to input a specific ticker and get performance data for that stock. Although not fully implemented at this stage, this additional feature would allow for even more flexibility, letting users dive deeper into specific stocks’ performance.

Step 7: Finalizing and Saving
After completing the tests and validating the results, I saved the workbook as a macro-enabled file (`.xlsm`) to ensure that the macro would be available for future use. The final product was a fully automated solution that could analyze stock data across multiple quarters with a single click.

Conclusion
The process of creating this VBA script allowed me to automate complex data analysis tasks, saving significant time and effort. By dynamically calculating and displaying the greatest percentage increase, decrease, and total volume across multiple worksheets, I was able to provide a clear and actionable view of stock performance. The addition of conditional formatting helped to visualize the data more effectively, making it easier to identify key trends at a glance.

This process not only simplified my workflow but also provided a solid foundation for future enhancements, such as dynamic user input for specific ticker analysis.
