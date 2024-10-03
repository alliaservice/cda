These scripts partially automate several steps in the library eBooks for courses purchasing workflow. These scripts are especially focused on automating the process of emailing faculty to let them know the library owns one or more eBooks for their course.

First a disclaimer: I learned python for this project, so there are lots of inefficiencies in my code and I'm sure there are many ways to improve the workflow. The entire project is a work in progress, and I will continue to improve it as I learn more. I'll outline my planned changes below.

# Summary

To understand the scripts, some background on how the library eBooks for courses purchasing program works is helpful. 

Every term UO instructors report all required books to our campus bookstore. The bookstore shares that booklist with the library, and we attempt to purchase eBook licenses for all the books on the list (that meet our criteria). To accomplish this goal we need to: 
1. Filter and sort the booklist to meet our selection criteria and to exclude any titles we've purchased or attempted to purchase in previous terms. (For books that we attempted to purchase but were unavailable we wait one year before searching again). The selection scripts help automate this process.
2. Use Alma to identify titles that are already in our catalog and to identify titles that we could potentially purchase (and where they are available). 
3. Manually check and purchase all available titles using our normal purchasing process. Catalog and activate the new titles using our normal cataloging process.
4. Notify faculty that one or more books from their courses is available as a library eBook. 
	-  I use both the notification scripts in this repository and a workflow in the Microsoft app Power Automate to do this. The scripts transform the data, and I use power automate as a more complex mail merge tool. I'll include details on how to set up a Power Automate workflow, but this process is contingent on using PA or another mail merge tool.

 The scripts in this repository help automate steps 1 and 4: the book selection process and the faculty notification process. I will include my workflow and documentation in the wiki, and a getting started guide for users who are not yet comfortable with python.
