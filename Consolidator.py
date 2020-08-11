
import os
import pandas as pd
from datetime import datetime

def main():
    #getting all of the files in this directory
    path = os.getcwd()+"//Inputs//"
    files = os.listdir(path)

    # the messege shown at the beginning of the script
    signal = input("Welcome to the Resume Score Consolidator. Please place scoring files from all reviewers in the “Inputs” folder,"+
    "and remove any unnecessary files from it. When you are ready, enter 1.")
    if signal == "1":
        # getting the current time--> used later for output files
        time_now = datetime.now()
        currentYear = time_now.year
        currentMonth = time_now.month
        currentDay = time_now.day

        #combining all of the datas in to one file and dropping all the unnecessary rows and cols
        combined_xlsx = [f for f in files if f[-4:]=='xlsx']
        all_data_ =pd.DataFrame()
        #the first ten rows are empty
        rows_to_skip = list(range(0,10))
        print('Currently reading all input files.')
        for f in combined_xlsx:
            data= pd.read_excel(path+f, skiprows= rows_to_skip)
            all_data_ = all_data_.append(data)
        print("File reading complete." + str(len(combined_xlsx))+" files ingested.")

        #drop the first columns and all the rows that are empty
        all_data_ = all_data_.drop(all_data_.columns[0],axis =1)
        all_data_=all_data_.dropna(axis=0,how='all')
        all_data_.fillna(0,inplace=True)
        all_data_.sort_values(by=['UID'],inplace=True)

        # create the final output dataframe
        final_output = pd.DataFrame()
        final_output=all_data_.copy()
        final_output = final_output.drop_duplicates(subset='UID',keep='first')
        final_output =final_output.drop(['Reviewer','Major','Athletics','Edu_Bonus','Work_Relevance','Work_Bonus','ECA','ECA_Bonus','CL','CL_Bonus','flag','comments','GPA ','Wtd_score'],axis=1)

        ave_score= [all_data_['Wtd_score'].iloc[[x,x+1]].mean() for x in range(0,len(all_data_),2)]
        final_output['Ave_Score_Final']= ave_score

        all_data_['flag_up'] =all_data_.flag.str.count('Yes')
        all_data_['flag_down']=all_data_.flag.str.count('No')
        flags_up=[all_data_['flag_up'].iloc[[x,x+1]].sum() for x in range(0,len(all_data_),2)]
        flags_down=[all_data_['flag_down'].iloc[[x,x+1]].sum() for x in range(0,len(all_data_),2)]
        final_output['Total_Flags_Up']= flags_up
        final_output['Total_Flags_Down']= flags_down
        final_output.sort_values(by=['Ave_Score_Final'],inplace=True)

        recon= pd.DataFrame()
        def greater(x):
            if (abs(all_data_['Wtd_score'].iloc[x]-all_data_['Wtd_score'].iloc[x+1])> 0.96):
                return True
        true= [greater(x) for x in range(0,len(all_data_),2)]
        if True in true:
            start =0
            for x in true:
                if x is True:
                    recon = recon.append(all_data_.iloc[2*true.index(x,start)])
                    recon = recon.append(all_data_.iloc[2*true.index(x,start)+1])
                    start =start+1

            recon = recon.drop(['Major','GPA ','Athletics','Edu_Bonus','Work_Relevance','Work_Bonus','ECA','ECA_Bonus','CL','CL_Bonus','flag_down','flag_up'],axis=1)
            recon= recon[['UID','First','Last','School','reviewer','Wtd_score','flag','comments']]
            print(str(int(len(recon)/2)) +' discrepancies are found. Please check Resume Reconciliation file.')
            recon.to_excel(os.getcwd()+"//Outputs//{}.{:02d}.{:02d} Resume Reconciliation.xlsx".format(currentYear, currentMonth, currentDay))

        # if no discrepancies found, then the consolidated resume screening file is created
        else:
            final_output.to_excel(os.getcwd()+"//Outputs//{}.{:02d}.{:02d} Consolidated Resume Screening.xlsx".format(currentYear, currentMonth, currentDay))

    # When the user has entered a wrong input
    else:
        print("You entered something other than 1. Aborting.")
        quit()



if __name__ == "__main__":
    main()
