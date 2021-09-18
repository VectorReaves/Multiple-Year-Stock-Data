# Multiple-Year-Stock-Data

The first step is the input external content as a variable. We do this is as "For Each ws in worksheet" For example.
Next step is to create the headers with strings using the reference of ws in ws.Range("I1").Value = "Ticker" and so forth for each one.
Variables: Create each variable as long or double depending on the intigers in the values.
Then we set the ticker using a for loop. totalTickerVolume = totalTickerVolume + ws.Cells(i, 7).Value means that variable in addition to the cells. "i" is 2 to lastrow and lastrow is a variable itself.
You create the if, elseif statments to determine the conditions.
