# Stock_Market_Analysis
  >Ticker_Tracker code outputs:
     
    -The name of tickers
    -The yearly change
    -the Percent change
    -The total stock volume
    -the greatest percent increase
    -the greatest percent decrease
    -the greatest overal stock volume
 
  >Possible Errors and Solutions:
    
    -Dividing by zero creating a percent change of infinity
        > To combat this error I imposed a if conditional that finds when the opening value = 0 and inputs an impossibly high value
        for the % change
          - That way the value can still be seen for greatest percent increase
        > Later I then impose another if conditional to find said impossibly high value and change it to a string as "Infinite"
        
    -Similarly There is also when both the Opening and Closing values are Zero
        >To combat this error I looked for which the sum of the opening and closing values equaled zero, and was able to input the
        percent change manually as zero through an if conditional.
