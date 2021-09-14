# -*- coding: utf-8 -*-
"""
Created on Sat Sep 11 18:09:55 2021

@author: favou
"""

from tabula import read_pdf
import pandas as pd
import numpy as np
import My_Functions as MF
from copy import deepcopy 
from Key_Function import append_df_to_excel


class TATA_Scrape():
    def __init__(self, page_list, write = 'No', Excel_file=None, header= "Yes", pdf_file="./ST.pdf"):
        
        self.page_list = page_list # list of pages to scrape e.g [16, 17]
        self.write = write 
        self.Excel_file = Excel_file #file where
        self.header = header
        self.pdf_file = pdf_file

    def raw_df_to_list(self):
        B = read_pdf(self.pdf_file, pages =self.page)[0]

        #raw headings 
        Headings = list(B.columns)

        C = B.values.tolist()
        C.insert(0, Headings) # inserts Headings list into the front of the list 

        #replacing 'nans' with 'x1' in list
        
        for i in range(len(C)):
            for j in range(len(C[i])):
                if isinstance(C[i][j], str) == False:
                    if isinstance(C[i][j], float):
                        C[i][j] = 'x1'
        
        self.C = C
        return B

    def row_numbers_begin(self):
        # determing the index where numbers begin in table C (number_index)
        j, H = 0, [] # row where we get numbers 
        while j < 5:        
            k =0
            k1 =0
            while k ==0 and k1 < len(self.C)-1:
                try:
                    L = float(self.C[k1][j])
                except:
        #            print("Oops!", sys.exc_info()[0], "occurred.")
                    l_remove=2 #remove this lines whenever
                else:
                    k =1
                k1 +=1
            H.append(k1-1)
            j +=1
        
        self.number_index = max(set(H), key=H.count)  # index of table, C, where numbers begin
       
        
    def cleaning_first_row(self):  
        # Cleaning the first row: unnamed ---> x1
        for i in range(len(self.C[0])):
            if 'Unnamed' in self.C[0][i]:
                self.C[0][i]= 'x1'
        
        # Cleaning the first row: Concatenating the strings to make headings
        for c in range(len(self.C[0])):
            J = []
            for r in range(self.number_index):
                if self.C[r][c] != 'x1':
                    J.append(self.C[r][c])
            if self.C[self.number_index][c] != "x1":
                self.C[self.number_index-1][c] = ""        
            for i1 in range(len(J)):
                self.C[self.number_index-1][c] = self.C[self.number_index-1][c] + " " +J[i1]
        
        self.D = self.C[(self.number_index-1):len(self.C)]
        
    def cleaning_serial_size_columns(self):
        ## Sorting out 'Serial size' column '
        
        # finding the column where 'Designation' resides 
        k =0
        while k ==0:
            for j in range(len(self.D)):
                i =0
                for element in self.D[0]:
                    if 'Designation' in element:
                        self.col = i
                        k =1
                    i +=1
    
        #Eliminating any double (or more) spacing in serial names  
        import re
        for i in range(len(self.D)):
            self.D[i][self.col] = re.sub(' +', ' ', self.D[i][self.col])        
    
        # Making the serial name column as a list
        SS = [self.D[i][self.col] for i in range(1,len(self.D))] 
        SS3 = MF.Serial_name(SS) # sorted column names as list
            
            
        
        # inputting the serial names into their appropriate rows in the main database, D
        i =0
        for element in SS3:
            i +=1
            self.D[i][self.col] = element        
                
            
            
            
    def cleaning_merged_columns(self):
        #finding columns that have been combined into one column by tabula 
        self.merged_col = []
        for i in range(len(self.D[1])):
            if i !=self.col and self.D[1][i] != 'x1':
                try:
                    float(self.D[1][i]) 
                except:
                    #print("Oops!", sys.exc_info()[0], "occurred.")
                    self.merged_col.append(i)
        
        
        #splitting columns that have been merged 
        E = deepcopy(self.D)
        e = 0 # ensures the current index is used as this changes with new columns added 
        z=[] # stores the index of the new columns that are created 
        for i in self.merged_col:
              for j in range(1,len(E)):
                  k = self.D[j][i].split(" ")
                  E[j][i+e] = k[0]
                  if i< len(E[0]) - 1:
                      E[j].insert(E[j].index(k[0])+1,k[1])
                  else:
                      E[j].append(k[1])
              z.append(E[j].index(k[0])+1)
              e +=1        
                  
        
        #cleaning merged columm headings (deleting any duplicated words)
        for i in self.merged_col:
            words = E[0][i].split() # splits the headings in question into a list of strings 
            E[0][i] = " ".join(sorted(set(words), key=words.index)) # removes the duplicated... 
            # ... from the list of strings then joins the items of the list inot one strings...
            #... and places this back into the heading 
        
        
        #Adding the new columm headings
        for i in z:
            old_head =E[0][i-1].split(" ")
            #break
            # if the heading has units 
            if MF.str_has_digit_m_N_or_k(old_head[len(old_head)-1]) == True:
                pen = old_head[len(old_head)-2]
                del old_head[len(old_head)-2]  #creates new head 
                del E[0][i-1]
                E[0].insert(i-1, " ".join(old_head) )
                old_head[len(old_head)-2] = pen
                E[0].insert(i, " ".join(old_head) )
            # if the heading doesn't have units
            else:
                end = old_head[len(old_head)-1]
                pen = old_head[len(old_head)-3]#
                del old_head[len(old_head)-1]
                del old_head[len(old_head)-2]# deleting web
                del E[0][i-1]
                E[0].insert(i-1, " ".join(old_head) )
                old_head[len(old_head)-2] = pen
                old_head[len(old_head)-1] = end
                E[0].insert(i, " ".join(old_head) )        
            
            
        # Finding empty columns( i.e. columns filled with 'x1')
          # these columns are have heading of 'x1'
        j =-1
        k = [] # stores the column that have heading 'x1'
        for i in E[0]:
            j +=1
            if i== 'x1':
                k.append(j)
        j =-1
        for element in k :
            j +=1
            for row in E:
                del row[element-j]
        self.E = E
    
    def final_df(self):
        # Converting list to dataframe
        df = pd.DataFrame([i for i in self.E[1:len(self.E)]], 
                            index = [i for i in range(1,len(self.E))],  
                            columns = self.E[0])   
        
        #moving 'Designation column if it appears as the final columns 
        col_headings = list(df.columns)
        if 'Designation' in col_headings[len(col_headings)-1]:
            df = df.iloc[:,0:(len(self.E[0])-1)]
        
            indicative_columns = []
            for i in list(df.columns):
                if 'ndicative' in i:
                    indicative_columns.append(i)
            
            for column_head in indicative_columns:
                df.drop(column_head, axis =1, inplace=True)
        
        return df
            
    def entire_process(self):
        
        df_store = []
        for page in self.page_list:
            self.page = page
            TATA_Scrape.raw_df_to_list(self)
            TATA_Scrape.row_numbers_begin(self)
            TATA_Scrape.cleaning_first_row(self)
            TATA_Scrape.cleaning_serial_size_columns(self)
            TATA_Scrape.cleaning_merged_columns(self)
            df = TATA_Scrape.final_df(self)
            df_store.append(df)
        
        df1 = pd.concat([df_store[0], df_store[1]], axis = 1)
             
        
        if self.write == "Yes":
            if self.header == "Yes":
                self.header = True
            else:
                self.header = False            
           
            append_df_to_excel(filename = self.Excel_file, df = df1, header = self.header, index = False)
        return df1
            

######################Testing ###############################################


pdf_file="./ST.pdf" # do not change unless you change the name of the source pdf
page_list = [16, 17]
write = 'No' # 'Yes or 'No'



#if writing
Excel_file = 'Test.xlsx'
header = 'No' # 'Yes' or 'No'


A = TATA_Scrape(page_list, write, Excel_file, header, pdf_file)         
B1 = A.entire_process()
B1
