#eval version#
## Declare Variable ##
MList = [] #List to collect value data
IList = [] #List to collect String value(formula) data
dic_val = {}
dic_value = {}
Mchoice = 0
Mrow = 0
Mcol = 0

## DEF ##

## input number ##
def inputNumber():
      global MList #Main MList
      global IList #Main IList
      global dic_val
      global dic_value
      s_col_row = str(input("Input cell like(row=1,col=a,input=3) A1 = 3 or A1 = 3 * 2, A1 = A1 * 2: "))
      s_col_row = s_col_row.split(' = ') #split cell by =
      formula = s_col_row[1] #s_col_row[1] is String value like 1, 10, A1 + A2
      cell = s_col_row[0] #s_col_row[0] is cell position like A1, B2
      col = ord(cell[0])-65 #col = number of col change from letter
      value = s_col_row[1] #split into first value, operator, second value
      row = int(cell[1:])-1 #check row from index 1 to last
      IList[row][col] = formula
      result = eval(value,dic_value) # calculate
      result = round(result,3) # round
      MList[row][col] = result
      dic_val[s_col_row[0]] = formula
      update_val() # update value

## choice ##
def choice():
      global Mchoice
      print("------------------ Welcome to Excel handmade ------------------")
      print("1. Creat your own table by input row and column")
      print("2. Print your table")
      print("3. Input your number in each cell in row and column like A1 = 3, A2 = A1 * 3 or else")
      print("q. To exit system")
      print("---------------------------------------------------------------")
      Mchoice = str(input("Input your command: "))
      print("---------------------------------------------------------------")
      select()

## select condition ##
def select():
      while Mchoice != 'q':
        if Mchoice == '1':
          creatTable()
          choice()
        if Mchoice == '2':
          print("\nThis is formula table")
          FormulaTable()
          print("\nThis is result table")
          Table()
          choice()
        if Mchoice == '3':
          inputNumber()
          choice()
        if Mchoice == 'q':
          break
        else:
            print("We have only 1, 2, 3 and q function please try again \n")
            choice()

##  creat table by column and row  ##
def creatTable():
    global IList 
    global MList
    global Mrow
    global Mcol
    global dic_value
    Mcol_char = input("Input column like A-Z: ")# input letter
    Mcol = ord(Mcol_char)-64 # change Mcol_char to number
    Mrow = int(input("Input row in range 1-100: ")) # input number
    
    if Mrow > 100: #check that Mrow or row cant more than 100
      print("Rows must less or equal 100")
      print('')
      return choice() # go to choice
    if Mcol > 26: #check that Mcol or column cant more than 26 or Z
      print("Columns must less or equal 26")
      print('')
      return choice() # go to choice
    for i in range(Mrow): # create 2D List (MList,IList) by rows and columns
        MList.append([''] * Mcol)
        IList.append([''] * Mcol)
    
    dic_value = {chr(65 + j) + str(i): 0 for i in range(1, Mrow + 1) for j in range(Mcol)}

## print table ##
def Table():
    global MList
    global Mrow
    global Mcol
    
    print("|    |", end='')
    # Print column headers
    for i in range(Mcol):
        print(f"  {chr(65+i)}  |", end='')
    print("\n" + "-" * ((Mcol * 6)+6))
    
    # Print cell values
    for j in range(Mrow):
        if j < 9:
          print(f"| {j+1}  |", end='')
        else:
          print(f"| {j+1} |", end='')
        for i in range(Mcol):
            cell_value = str(MList[j][i])
            padding = " " * (3 - len(cell_value))
            print(f" {padding}{cell_value} |", end='')
        print("\n" + "-" * ((Mcol * 6)+6))

def FormulaTable():
    global IList
    global Mrow
    global Mcol
    
    print("|    |", end='')
    # Print column headers
    for i in range(Mcol):
        print(f"  {chr(65+i)}  |", end='')
    print("\n" + "-" * ((Mcol * 6)+6))
    
    # Print cell values
    for j in range(Mrow):
        if j < 9:
          print(f"| {j+1}  |", end='')
        else:
          print(f"| {j+1} |", end='')
        for i in range(Mcol):
            cell_value = str(IList[j][i])
            padding = " " * (3 - len(cell_value))
            print(f" {padding}{cell_value} |", end='')
        print("\n" + "-" * ((Mcol * 6)+6))

def update_val():
    global MList
    global IList
    global dic_val
    global dic_value
    global Mrow
    global Mcol
    for k,v in dic_val.items(): # loop check value from dic_val
      result = eval(v,dic_value) # calculate using dic_value
      cell = k # cell position by key of dic_val
      col = ord(cell[0])-65 #col = number of col change from letter
      row = int(cell[1:])-1 # row = index 1 to last
      MList[row][col] = result
      dic_value[k] = result
    
  
## MAIN ##
choice()