
#Feed 2 parallel lists to this and it will return a dictionary.
def createdict():

    ##GET DATA F
    goodFile = False
    while goodFile == False:
        #calls to Dialog.py and opens a dialog box for the user to select a file. 
        fname = getfiledir()
        # Begin exception handling
        try:
            # Try to open the file using the name given
            olympicsFile = open(fname, 'r')
            # If the name is valid, set Boolean to true to exit loop
            goodFile = True
        except:
            # If the name is not valid - IOError exception is raised
            print("Invalid filename, please try again ... ")


    List1 = []
    List2 = []
    next(olympicsFile)
    #loop through file line by line and write to lists, sku, item, qty are the temp variables, SKU, Qty, Item are the lists. 
    for line in olympicsFile:
        line = line.strip()
        x, y = line.split(',')
        List1.append(x)
        List2.append(y)

        
    olympicsFile.close

    ##CREATE DICT STARTS
    lookupObj = zip(List1, List2)
    Dict1 = dict(LookupObj)
    print (Dict1)
    


    return Dict1







    
    
