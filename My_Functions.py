# -*- coding: utf-8 -*-
"""
Created on Tue Jun 22 13:02:29 2021
My funtions

@author: favou
"""
import pandas as pd
from math import isnan
def nan_DF(DF):
    """"This function takes in a pandas dataframe coloumn or list, changes its name to 0
    and determines if it is filled with 'Nan'"""
    import pandas as pd
    c = pd.DataFrame(DF) #converts input into a pandas DF if not already a DF
    x = c.columns[0]# determines the heading of the DF(which corressponds to index 0 i.e. column heading)
    c = c.rename(columns={x: 0}) # changes the column heading to 0
    d = pd.DataFrame([float('nan')]*len(c)) #creates a DF column of nan values with the 
    if c[0].equals(d[0]):
        return True
    else:
        return False
    
#alternatively 
#covert database cloumn to list: c = list(DF[0]) then use the following
#funtion to determine if the list is all nan:
    # np.isnan(c).all()
#This should return a boolean 

def List_mode(L):
    """" This function takes in a list of integers, L, and determines the most frequent
    integer of the list (i.e. the mode) as an integer"""
    element_count = [] # this stores items of list, L, and the amount of times they appear in L
    k =0 # this is the index of the 1st item of the list
    a=len(L) # this is the number or items in the list
    
    while a >0: # iterate until there is no more items in the list 
        count =0 # counts the number of times an item appears in the list
        
        for i in range(a): # iterates through the list to find the number of times the...
            if L[k] == L[i]: #...1st element of the list apears in the list,L 
                count += 1
        element_count.append(str(L[k])) # inputs the items of list, L as a string
        element_count.append(count) # inputs the amount of times that item appears in list, L
        L = list(filter((L[k]).__ne__, L)) # removes all instances of the an item L[k] in list L
        a = len(L) # recalcuates the lenght of the list(now that all instances of the 1st item have been removed)
        #print(L)
    count_list =[] # this stores all the counts for all items in the list
    # finding actual mode
    
    for i in range(1, len(element_count), 2): # iterates through index of list to...
        count_list.append(element_count[i])  #... append all counts of the items of list L (ignoring the actual item )
    mode_frequency = max(count_list) # determines the max count of the list
    index_mf = element_count.index(mode_frequency)  #...determine index of max count in element_count
    mode = int(element_count[index_mf-1]) # determines the mode 
    return(mode)
#alternatively 
#just use: max(set(L), key=L.count)                
            
def Serial_name(SS):
    """" This function takes the serial name column (a list of strings, SS,) 
         joins the elements of the list into a single string, splits them by
         spacing then determines the appropriate serial name in form <<AxBxC>> """
    # joins the list strings into one string in a list
    SS1 = [' '.join(SS)] 
    #splits the string in the list SS1 into a list of strings seperated by a space
    SS2 = [i.split(" ") for i in SS1][0]
    # #creates a list of the full serial names (i.e. '1016x305x487')
    SS3 = []
    j =-1
    k, t =0,0
    p=0
    while j< len(SS2)-1: #examines all elements in SS2
        j +=1
        i = SS2[j]
        if 'x' in i or 'A' in i: #examines the elements in SS2 that have 'x'
            k =j # indicates the index for the element that contain 'x' in SS2
            t =1 if k ==0 else None
        else: #creates a new list where all serial names are inputted without any unnecessary rows
             if j ==0 and t ==0:
                 p=1
                 k =1
             if p==1 and j != len(SS2)-1 and j !=0:
                if 'x' in SS2[j+1]:
                    k = j
                    SS2[k],SS2[k+1] = SS2[k+1], SS2[k] # that have just the begining of the serial names
                    j = j+1
            
             SS3.append(SS2[k]+"x"+SS2[j]) 
    return(SS3)


def str_has_digit_m_N_or_k(S):
    any([char.isdigit() for char in S])
    k =0
    for char in S:
        if char.isdigit() == True or char == 'm' or char == 'N' or char == 'k':
            k = 1
    if k ==1 :
        return True 
    else: 
        return False 

def atlest_one_char(S):
    k =0
    for char in S:
        if str(char)!= ' ':
            if str(char).isdigit() == False:
                k = 1
    if k ==1 :
        return True 
    else: 
        return False     
    
def underline(string):
    a = '\033[4m'
    b = '\033[0m'
    return a+string+b


def merged_column(D, row = None, s = 'Y'):
    """ This function find the columns that have been merged by tabular based
        on whether or not the mid row contains digits sparated by space"""
    if isinstance(D, list) == False:
       E = D.values.tolist()
       D = E
    if s == 'Y':
        row = int(round(0.5*len(D)+1, 0))
    merged_col = []
    for i in range(len(D[row])):
        if isinstance(D[row][i], float) == False:
            if atlest_one_char(D[row][i]) == False:
                try:
                    float(D[row][i])
                except:
                    #print("Oops!", sys.exc_info()[0], "occurred.")
                    merged_col.append(i)
    return merged_col

def digit_spitter(D, numbers_start=0):
    from copy import deepcopy
    merged_col = merged_column(D)
    E = D.values.tolist()
    D = deepcopy(E)
    
    for i in merged_col:            
        for j in range(numbers_start,len(E)):
            k = D[j][i].split(" ")
            del E[j][i]
            for n in range(len(k)):
                E[j].insert(i+n,k[n])    
    return E

def raw_table(file, page, csv_file, numbers_start=0):
    """ file = the file where the raw table is ectracted from, 
        page = the page where the table is extracted from
        csv_file = the filename for the output csv file: 'filename.csv'
        number_start = the row index where the numbers start. this is achived
                       iteratively (i.e. you need to run the function once to determine this). 
    """
    from tabula import read_pdf
    raw_table = read_pdf(file, pages =page)[0]
    refined = digit_spitter(numbers_start = numbers_start, D=raw_table)
    
    import csv

    with open(csv_file, "w", newline="", encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(refined)
    return refined


def searching_df_columns(df, column_name):
    """This searches the column contaning the strings in column_name and 
       returns the column name as a string """
    found =0
    while found ==0:
        for col in df:
            if column_name in col:
                found =1
                break
    return col

def member_data_locator(beam_column, serial_number, column_heading):
    """ Takes in the tyoe of member (beam or column), its serial number and some 
        words in its heading (doesn't have to be all) and extracts the relevant 
        property for a beam/column for that serial number """
        
    UKC_dir = 'C:/Users/favou/Desktop/Summer Python/Merged UKC.xlsx'
    UKB_dir = 'C:/Users/favou/Desktop/Summer Python/Merged UKB.xlsx'
    from pandas import read_excel 
    if beam_column == 'column':
        df = read_excel (UKC_dir)
        
    elif beam_column == 'beam':
        df = read_excel (UKB_dir)               

    data_col = searching_df_columns(df, column_heading)
    row = df.loc[df['Designation Serial size'] == serial_number]
    row.index = [0] #changes the row index to the integer 0  
    data = row.loc[0, data_col]
    return data

#AAA = raw_table('Nominal_thickness.pdf', 32, 'Steel_grade1.csv', 2)

def any_True(value, lst):
    
    """ This function returns a True boolean if any item in the list(lst) is 
        equivalent to value """
        
    a = [] 
    for i in lst:
       if value == i:
           a.append(True)
       else:
           a.append(False)
    
    return any(a)
       