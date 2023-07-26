# AutomateExcel

Automates Taarcom's Excel processes for insight reports from our distributors

Insight reports (input) come in the form of Excel spreadsheets, containing 
buy orders (new business opportunities) from potential customers. The process of
extracting useful information from these reports includes (1) standardizing the
data, (2) splitting the data into separate files for each sales rep, and (3) 
recieving the files back from each sales rep with their feedback and compiling
all cleaned, commented insight files back into one master file.

Executes three tasks: clean, split, and compile

Uses the following lookup tables to complete its tasks:
    - Master Account List (define each sales rep's accounts)
    - Zip Code Territory List (define each sales rep's territories)
    - Root Customer Mapping (map reported customer names to actual customers stored on account list)
    - Report Columns (map input file column headers to desired standard column headers)

Dumps everything into a single output file.

## Clean
      Maps input file columns to their desired column headers
      Maps reported customer to actual end customer stored in the account list
          if not in the account list, highlights yellow (new business!)
      Assigns each account to a sales rep
          first, checks if the account is on the account list
          then, checks by territory

## Split
      Creates a new insight file for each sales rep, containing only their line items

## Compile
      Stitches together any amount of Excel files, as long as they have identical columns
      Used to compile sales rep feedback files


