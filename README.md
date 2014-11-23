### Excel --> Database

## What 

I created this tool at work to help analyze information "contractors" modified. The "contractors" had 1,000+ workbooks where they reviewed rates and volumes. After their work was finished, I needed to crawl through what they did and store it in a tabular format, that would then be uploaded to a database.

## How

The VBA code iteratively opens all files in the path you give it, and stores the contents in an array variable. 

It then switches to the main `2database` file, and inserts the contents from the array variable.

Loop through this cycle until there are no more records to insert.

Also, the information is grouped by year, so 1 excel file <> 1 record... some clever trickery to get this to work.

Enjoy!