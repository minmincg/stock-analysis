# stock_analysis summary

The first thing I did was to iterate through all of my worksheets using a for each WS
<img width="390" alt="Screenshot 2023-01-15 at 19 59 18" src="https://user-images.githubusercontent.com/64171430/212582839-1ccaba35-5e75-4f7b-bd39-0ac95f2b13c3.png">


Then I iterated through the tickers and made an If conditional where if my ticker changed, I stored the current ticker's last day's closing value and stored the NEXT ticker's first day opening value. 

<img width="555" alt="Screenshot 2023-01-15 at 20 00 40" src="https://user-images.githubusercontent.com/64171430/212582952-f1970c1d-a351-42c6-9d9d-7f490b992cea.png">


I did a subtraction (closing value - opening value) to get the difference 

I declared the new values to a new column were I wrote the ticker in question, the yearly change in dollars, the percentage change, and the total stock volume for each ticker in that year. 
<img width="223" alt="Screenshot 2023-01-15 at 20 01 59" src="https://user-images.githubusercontent.com/64171430/212583053-b56844a0-d56a-4716-891c-c2ff6c84e15e.png">

Then, I did another iteration to go through the new menu to discover which ticker had the most percentage increase, and which one had the most negative increase and added their value. Also I summed up the total stock volume for all of the tickers.

<img width="179" alt="Screenshot 2023-01-15 at 20 03 58" src="https://user-images.githubusercontent.com/64171430/212583230-a1e3dee5-fc99-4f08-87a8-92630a26efe1.png">

Finally I made sure it worked on my original data since I was using a test version so the script would run faster while developing. 

Here are the final results of the code in my original data. 
![stock_results_image](results.png)
