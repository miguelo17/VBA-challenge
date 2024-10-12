sub Analyzer ()

    'declare variables for application
    Dim i As Integer
    Dim j As Integer
    Dim ticker as String
    Dim qchange as Double
    Dim pchange as Double
    Dim svolume as Double
    
    'initialyze values for variables
    qchange = 0
    pchange = 0
    svolume = 0

    Dim openingtotal as Double
    openingtotal=0

    Dim closingtotal as Double
    closingtotal=0

    dim summaryrowtable as integer
    summaryrowtable=2

    'find last row in sheet A
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
  

    'loop through all rows in column 1

    for i =2 to lastRow

        'check if we are on the same ticker name
        if cells(i+1,1).Value <> Cells(i, 1).Value then

        'set ticker name
         ticker= cells(i,1).Value

        'add to the opening total
         openingtotal = openingtotal + cells(i,3).value

        'add to the closing total

         closingtotal = closingtotal + cells(i,6).value

        'total stocks volume
         svolume = svolume + cells(i,7).value

        'quaterly change operation
         qchange =  closingtotal - openingtotal

        'percent change operation
         pchange = ((closingtotal / openingtotal))*100

        
        'print summary table headers
        range ("I1") = "Ticker"
       
        range ("J1") = "Quaterly Change"
       
        range ("K1") = "Percent Change %"
        
        range ("L1") = "Total Stock Volume"
      
        
        'print summary table data

        range("I" & summaryrowtable).value = ticker

        range("J" & summaryrowtable).value = qchange

        range("K" & summaryrowtable).value = pchange

        range("L" & summaryrowtable).value = svolume
      

        'add a row to the summary table for the next item

        summaryrowtable = summaryrowtable +1

        'reset variables for the next items

        qchange = 0
        pchange = 0
        svolume = 0

        else 
            svolume = svolume + cells(i,7).value
            openingtotal = openingtotal + cells(i,3).value
            closingtotal = closingtotal + cells(i,6).value
            
        end if

    next i

    '***Finding Greater Increase ***'

    'Search for last row in second summary table
    Dim lastRow2 As Long
    lastRow2 = Cells(Rows.Count, "K").End(xlUp).Row

    Dim maxvalue as Double
        maxvalue = 0


    Dim maxstocks as Double
        maxstocks = 0

    'loop through all row to search the greater Value for Increase
    for i = 2 to lastRow2

        if cells(i+1,11).Value > Cells(i, 11).Value then

            maxvalue= cells(i+1,11)

            elseif cells(i+1,11).Value < Cells(i, 11).Value then

            maxvalue= cells(i,11)
            

        
        end if

    next i
    'loop through all row to search the greater Value for Stocks
    for i = 2 to lastRow2

        if cells(i+1,12).Value > Cells(i, 12).Value then

            maxstocks= cells(i+1,12)

            elseif cells(i+1,12).Value < Cells(i,12).Value then

            maxstocks= cells(i,12)
                 
        end if

    next i

    range ("N1") = "Ticker"

    range ("O1") = "Greater Increase"
       
    range ("P1") = "Greater Stock Value"

    range("N2")= ticker
    range("O2")=maxvalue
    range("P2")=maxstocks

end sub

'*******trial subroutine to find max values on the list*******'
'sub Max ()
    'last row of summary table
   '   Dim lastRow2 As Long
    'lastRow2 = Cells(Rows.Count, "B").End(xlUp).Row

   ' Dim maxvalue as Double
   '     maxvalue = 0

   ' range("C" & lastRow2)=lastRow2

  '  for i = 2 to lastRow2

      '  if cells(i+1,2).Value > Cells(i, 2).Value then

       '     maxvalue= cells(i+1,2)

       '    elseif cells(i+1,2).Value < Cells(i, 2).Value then

       '     maxvalue= cells(i,2)

        
     '   end if

  '  next i

   ' range("C2")= maxvalue

'end sub


