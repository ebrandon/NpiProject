# NpiProject

While consulting, I had a client that required data cleanup for some government reporting obligations. Sales reps of pharma companies often have meals with healthcare professionals and the value of these meals has to be reported to the Centers for Medicare and Medicaid Services. This is a fairly new requirement and the year I did it for the client was the first year in which reporting was required. Each healthcare professional has a license number that is used to uniquely identify them, called an NPI number (National Provider Identifier). However, since the requirement was so new, most pharma companies did not plan for how complicated it would be to track down the information for each physician and they did not require that sales reps input this information into the expense systems when reporting their expenses. The NPI numbers therefore had to be manually looked up in order to report the information CMS. There were many problems such as physician addresses not being updated in the government's system, many physicians with the same name practicing within the same city/state, unclean data entered by the sales reps in the expense reporting systems, etc. 

The client gave me a spreadsheet with tens of thousands of rows of data and showed me a website where you could enter a physician's name, address or other information and the website would show a list of potential NPI matches. They asked me to go through each row of data and come up with matches for each of the physicians, working with the sales reps when there were multiple possible matches to resolve the issues. The website was very slow and also required the user to input a captcha code for each search performed.

I decided that if I did this work manually, it would take longer than the amount of time we had before the deadline to report to CMS. I instead found an open API to the government's data called BloomAPI and built a python app to pull in all possible matches based on the physician's name. The client had multiple different formats of spreadsheet coming from different expense systems. So I built a simple interface using tkinter so that I could specify which columns to use in the input spreadsheets as the input columns. I also included a separate very similar app to update the physician's address and other information for physicians in which the client already had a matching NPI number so that the address we reported to CMS would match the address that CMS had on file. 

The program takes in a spreadsheet, and allows you to specify the columns to use as input columns. While running it shows how many matches were obtained for each physician name and outputs "Finished" to the interface once complete. The output is a second spreadsheet that shows all of the original lines of data highlighted in yellow with possible matches listed on the rows directly underneath each yellow line of original data. I was then able to use those spreadsheets as a starting point to manually choose the correct NPI number based on online research about the physicians and conversations with sales reps. 

# Requirements
The program runs using Python 3. Requirements are included in the requirements.txt file.

To run the program, go to the python folder and use the command - python3 npiProgram.py

There is a sample file in the data folder that can be used as an input file

![npi screenshot](https://user-images.githubusercontent.com/3095171/28758434-94bc3448-754c-11e7-83b7-f6fb141bfffd.png)



