import sys
import pandas as pd
import json
import pickle
import os
import time


def validate_filepath(filename):
    if not os.path.exists(filename):
        raise FileNotFoundError("Filepath : {} does not exists !!".format(filename))
        
def read_excel(filepath):
    """
    Return the dataframe of the excel file.
    Arguments:
        filepath: excel file path.
    Returns : Dataframe
    """
    validate_filepath(filepath)
    df = pd.read_excel(filepath)
    df = df.astype(str)
    return df

def add_rows(addfile, filepath):
    """
    This function add the rows to the existing file.
    Arguments:
        addfile : path of the comma separated file that contains rows to be added.
        filepath : the excel file to which rows need to be added.
    Returns : Excel file
    """
    employeeTable = read_excel(filepath)

    validate_filepath(addfile)
    add_rows = []
    with open(addfile, "r") as file:
        for line in file.readlines():
            add_rows.append(line.rstrip().split(","))
    for element in add_rows:
        if element[0] in employeeTable['Employee Id'].tolist():
            print("ADD ROWS FAILED !! /nEmployee Id : {} already exists !!".format(element[0]))
            continue
        else:
            employeeTable.loc[len(employeeTable)] = element 
    print(employeeTable)
   

def delete_rows(delfile, filepath):
    """
    This function delete the rows from the file based on the Employee ID.
    Arguments:
        delfile : Path of the comma separated file that contins Employee ID.
        file path : the excel file path from where the rows need to be deleted.
    Returns : excel file
    """
    employeeTable = read_excel(filepath)
    del_rows = []
    validate_filepath(delfile)
    with open(delfile, "r") as file:
        for line in file.readlines():
            del_rows.append(line.rstrip())
    for i in del_rows:
        if i in employeeTable["Employee Id"].tolist():
            employeeTable = employeeTable.drop(employeeTable[employeeTable["Employee Id"] == i].index).reset_index(drop=True)
        else:
            print("DELETE ROWS FAILED !! /nEmployee Id : {} doest not exists !!".format(i))
    print(employeeTable)
    
    
def update_rows(updatefile, filepath):
    """
    This function update the rows based on the Employee ID.
    Arguments:
        updatefile: path to the comma separated file that contains dictionary values.
        filepath: path to the excel file to which the rows need to be updated.
    Returns: Excel file
    """
    employeeTable = read_excel(filepath)
    validate_filepath(updatefile)
    update_rows = []
    with open(updatefile,"r") as file:
        for line in file.readlines():
            update_rows.append(json.loads(line))
    print(update_rows)
     
    for element in update_rows:
        if element["Employee Id"] not in employeeTable["Employee Id"].tolist():
            print("Employee Id : {} does not exists !!".format(element["Employee Id"]))
        else:
            if "Employee Name" in element.keys():
                employeeTable.at[employeeTable[employeeTable["Employee Id"] == element["Employee Id"]].index,"Employee Name"] = element["Employee Name"]
            if "Year of Joining" in element.keys():
                employeeTable.at[employeeTable[employeeTable["Employee Id"] == element["Employee Id"]].index,"Year of Joining"] = element["Year of Joining"]
    print(employeeTable)
    
                
if __name__ == "__main__":
    if not os.path.exists(sys.argv[1]):
        raise FileNotFoundError("Filepath : {} does not exists !!".format(sys.argv[1]))
    if len(sys.argv) > 2:
        if sys.argv[2] == 'add':
            add_rows(addfile=sys.argv[3], filepath=sys.argv[1])
        if sys.argv[2] == 'del':
            delete_rows(delfile=sys.argv[3], filepath=sys.argv[1])
        if sys.argv[2] == 'update':
            update_rows(updatefile=sys.argv[3], filepath=sys.argv[1])
    else:
        print(read_excel(sys.argv[1]))