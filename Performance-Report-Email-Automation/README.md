# Automatising the Creation and Sending of Email Campaign Performance Report using GoogleScript

I created a script to automate the email report that I was required to write and send to management team every few days.

I had previously created a solution to generate and aggregate all client campaign data into one google sheet to create a campaign overview. With an aim to reduce duplication of work and time spent writing emails by hand, I tasked myself with writing a script despite having no previous Javascript experience.

This script effectively worked as 'mail merge' pulling information from a specific cells from the main sheet and inserting these pieces of data into a piece of html code I had written in another sheet in the same workbook and then having to send it to whoever I wish by entering the email addresses into a specific cell.

I designed this solution to be extensible allowing the contents of the email to be changed to what ever the user wants.
