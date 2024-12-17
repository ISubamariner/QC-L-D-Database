#File Merger for QC SHRU L&D Database, Created by Ian Salig U Batangan, contact details:isubatangan@gmail.com
# Purpose: to automatically merge all the training data available

def pip(): #install needed 3rd party libraries
    import sys
    import subprocess
    # implement pip as a subprocess:
    packagename=('pandas','pathlib','openpyxl')
    for i in packagename:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', i])
    # process output with an API in the subprocess module:
    reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'freeze'])
    installed_packages = [r.decode().split('==')[0] for r in reqs.split()]
    print(installed_packages)
#pip()

#for data manipulation library, also needs openpyxl sub library of pandas
import pandas as pd

#string manipulation library
import re

#for finding file path library
import glob
from pathlib import Path


def merge(variable_name,root_path):

    #determining path
    path_reg = f"{root_path}/{variable_name}_Reg.xlsx"
    path_post = f"{root_path}/{variable_name}_Post.csv"

    #reads raw excel and csv
    df_reg = pd.read_excel(path_reg)
    df_post = pd.read_csv(path_post)
    
    df_reg.columns = map(str.upper, df_reg.columns)
    df_post.columns = map(str.upper, df_post.columns)
        
    #cleaning of column names
        #creating a list of the column names
    df_reg_columns = df_reg.columns
    df_post_columns = df_post.columns 
        
        #checking the list if the data set has assessment
    if 'PRE-ASSESSMENT TOTAL' not in df_reg_columns or 'QTOTAL' not in df_post_columns:check_pre=1
    else: check_pre= 0

        #mass removal of extrenous data
                            #ISUB
    df_reg_cleaned = [e for e in df_reg_columns if "PRE-ASSESSMENT" not in e and 'EMAIL ADDRESS' not in e and  \
                          'NICKNAME' not in e and 'ENDORSEMENT LETTER' not in e and 'CSC UPLOADED' not in e and 'DATE ANSWERED' not in e and 'EXPECTED OUTCOMES'not in e and 'DATA PRIVACY CONSENT' not in e and 'CONTACT NUMBER' not in e]
        
            #using re library to exlude columns with Q in name in specific cases
    df_post_cleaned = [e for e in df_post_columns if not re.match(re.compile('Q.+-' ) , e) and not re.match(re.compile('Q..' ) , e)and not re.match(re.compile('Q.' ) , e) \
                        and 'EMAIL ADDRESS' not in e and 'NICKNAME' not in e and \
                        'ENDORSEMENT LETTER' not in e and 'CSC UPLOADED' not in e and 'DATE ANSWERED' not in e and \
                            'EXPECTED OUTCOMES'not in e and 'DATA PRIVACY CONSENT'not in e and 'CONTACT NUMBER' not in e]
        
        #adds back the needed column since mass deletion was done previously
    if check_pre == 0:
         df_reg_cleaned = df_reg_cleaned+['PRE-ASSESSMENT TOTAL']
         df_post_cleaned = df_post_cleaned+ ['QTOTAL']
        
        
    #finding maximum score
    Qmax_list = [e for e in df_post_columns if re.match(re.compile('Q.+-' ) , e) ]
    max_score=len(Qmax_list)
    

    #retrieving specific column names
    df_reg = df_reg.loc[:,df_reg_cleaned]
    df_post = df_post.loc[:,df_post_cleaned]
    df_post= df_post.assign(Maximum_Assesment_Score=max_score,Training_Code=variable_name)

    #debugging
    '''
    print('\n this is registration column \n',df_reg_cleaned,'\n this is post columns\n',df_post_cleaned,'\n this is post data shape \n',df_post_shape, '\n')    
    '''
    #Changes all data in dataframe as string
    for name in df_reg.columns:
        df_reg[f'{name}'] = df_reg[f'{name}'].astype(str)
    
    for name in df_post.columns:
        df_post[f'{name}'] = df_post[f'{name}'].astype(str)

        #Left Join of post and reg data and drops duplicates
        #checking if data set has full name if not full name is created
    if 'FULL NAME' in df_post_columns and 'FULL NAME' in df_reg_columns:
        df_merge= df_reg.merge(df_post,left_on=['FULL NAME',"DESIGNATION/POSITION","DIVISION/ SECTION",\
                                                    'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX'],\
                                                        right_on=['FULL NAME',"DESIGNATION/POSITION","DIVISION/ SECTION",\
                                                    'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX']).drop_duplicates()
    else:
        df_merge= df_reg.merge(df_post,left_on=["LAST NAME","FIRST NAME","MIDDLE INITIAL","DESIGNATION/POSITION",\
                                                    "DIVISION/ SECTION",'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX'],\
                                                            right_on=["LAST NAME","FIRST NAME","MIDDLE INITIAL","DESIGNATION/POSITION","DIVISION/ SECTION",\
                                                        'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX']).drop_duplicates()
        #generate Full Name
        df_merge['FULL NAME'] = df_merge["FIRST NAME"]+ " " + df_merge["MIDDLE INITIAL"] + " "+ df_merge["LAST NAME"]
        
    #saves the merged data to a csv file
        
    df_merge.to_csv(f"{root_path}/MergedFiles/{variable_name}_merged.csv")
    return df_merge
    

def main():
    #finding file path
    root_path=Path.cwd()
    print(f'File path is {root_path}')

    #checks current folder for all the trainings with post assesment and makes a list of the unique trainings
    path_list=glob.glob(f"*_Post.csv")
    training_list=[]
    main_df=pd.DataFrame()
    for name in path_list:
        training_list.append(name[:-9:])
    total_files=len(training_list)
    

    #uses list of trainings to look for the file to merge
    i=0
    for name in training_list:
        i+=1
        temp_df=merge(name,root_path) #temporary dataframe
        main_df=pd.concat([main_df,temp_df]) #concatenates the data
        print(f'{name} is merged, with a shape {temp_df.shape} \n {i} out of {total_files}')
    
    #Re-arranging the Data Columns
    main_df_columns= main_df.columns
        #main information
    main_df_columns_tag=["ID",'Training_Code',"FULL NAME","LAST NAME","FIRST NAME","MIDDLE INITIAL",\
                          "DESIGNATION/POSITION","DIVISION/ SECTION",'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX','PRE-ASSESSMENT TOTAL', 'QTOTAL', 'Maximum_Assesment_Score']
    
    main_df_columns_design=[e for e in main_df_columns if re.match(re.compile('PROGRAM DESIGN.+' ) , e) ]
    
    main_df_columns_trainingmat=[e for e in main_df_columns if re.match(re.compile('TRAINING.+' ) , e)]

    main_df_columns_logistics=[e for e in main_df_columns if re.match(re.compile('LOGISTICS.+' ) , e)]

    main_df_columns_expectations=[e for e in main_df_columns if re.match(re.compile('EXPECTATION.+' ) , e)]
    
    main_df_columns_administration=[e for e in main_df_columns if re.match(re.compile('ADMINISTRATION.+' ) , e)]

    main_df_columns_comments=[e for e in main_df_columns if re.match(re.compile('COMMENT.+' ) , e)]

    main_df_columns_facilitators=[e for e in main_df_columns if re.match(re.compile('FACILITATOR.+' ) , e)]

    main_df_ordered_list= main_df_columns_tag + main_df_columns_design + main_df_columns_trainingmat + main_df_columns_logistics + main_df_columns_expectations + main_df_columns_administration + main_df_columns_comments + main_df_columns_facilitators
    
    main_df_columns_others=[e for e in main_df_columns if e not in main_df_ordered_list]

    main_df_ordered_list= main_df_columns_tag + main_df_columns_design + main_df_columns_trainingmat + main_df_columns_logistics + main_df_columns_expectations + main_df_columns_administration + main_df_columns_comments + main_df_columns_facilitators + main_df_columns_others

    #ISUB#
    main_df= main_df[main_df_ordered_list]

    print((main_df_ordered_list+main_df_columns_others))

    #feedbacka
    print('Main DataFrame is Updated')
    print(f'The total data shape is {main_df.shape}')
    main_df.to_csv(f"{root_path}/MergedFiles/AllConcat.csv") 


    main_df.describe()


#calls the code
main()