import pyodbc 
from logging import root
from sys import exit
import tkinter as tk 
import pyodbc


# Backend pyodbc starts here

def show_entries():
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\2022000130\Downloads\Company_Data.accdb;') 
    
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM Company_data1')

    for row in cursor.fetchall():
        print (row) 
     
def quit_program():
    '''Exits the program, including the frontend and backend'''
    root.destroy()
    exit()   

root = tk.Tk()

def __init__(self): 
    pass 

def print_all(self): 
    print("showing all records") 
    

def pos_growth(self): 
    print("all + growth") 