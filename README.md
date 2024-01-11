# vba-challenge
# The file "Module_2_code.bas" is a file containing VBA code that works on both Multie_Year_Stock and Alphabetical_testing. The code does the following:
# Loops through both the Alphabet and Multi_Year_Stock_Data worksheet and does the following:
# Posts every unique ticker.
# Adds together all the volume values for each ticker
# Calculates the yearly change of each ticker by taking the closing price of the ticker on it's last date and subtracting the opening price on the ticker's opening date.
# Calculates the percent change of each ticker by taking the yearly change and dividing it by the opening price
# Loops through these entries and finds if the yearly change was an increase or a decrease and color codes it to green or red (respectively).
# Loops through and finds the ticker with the greatest increase, the ticker with the greatest decrease and the ticker with the greatest total volume and posts them to a blank part of the Excel spreadsheet with the ticker name and the value assosciated with it.
# While I was looking for inspiration for my code, I had found some code that answered a similar problem. From the line TotalVolume = TotalVolume + ws.Cells(i, 7).Value()" to "Next i" I had gotten inspiriation from this link:
# https://github.com/shrawantee/VBA-Scripting---Stock-Market-Analysis/blob/master/HW2_Moderate_DS.vbs
