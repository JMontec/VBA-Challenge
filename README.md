# VBA-Challenge
First from the instructions I understood I needed to create loops for the ticker, yearly change, percent change, and stock volume
since ticker was all words I set it as string
since all the sheets in the excel ended at different rows I set up a LastRow function
Using examples from previous assingments I knew I would have to make a function to list out all the tickers and the corresponding values
Since these were going to be input I set my ranges
I started with my For loop to start at 2 since row 1 was the title
My first if statement was to catch that first tickers opening amount followed by my second if statment to get the closing amount
Having now grabbed these values I make my functions to the corresponding values the assignment asked for the yearly change, and percent change
I also include the volume count as it was reading down the ticker to tally the sum of the volume until hitting a new ticker symbol
I did not see it specifically mention that we needed to color code the yearly change but since it was on the example I made a quick function
I set it up for the range to follow down the yearly change column if it was more than zero it would fill in green if less it would fill in red
Next I input functions to fill in the columns of I, J, K, and L
I also made sure that each row would have the correct values and worked on the second part of the assignment
Now that I had the column set up for the changes yearly and percent
I created a new VBA code for comparing the results and finding the highest % increase, lowest % decrease, and greatest total volume
I created the range and values to set the output for my functions
I set the function for the values I am trying to find
starting my For at 2 since row one is the title
I set up my functions to search each row seeing if it is more than zero and if it was to hold that number going down until the next row it finds is more than that current row its comparing to
The same was set for the opposite searching for less than zero and the biggest negative number
For volume it was a bit easier since I just had to set a function to compare the biggest numbers going down the volume column
