#%% Libraries
import requests
import pandas as pd
from collections import Counter
import matplotlib.pyplot as plt
import os
import glob
import tkinter as tk
import sys
from matplotlib.backends.backend_pdf import PdfPages
from mtranslate import translate
import pickle
try :
    with open('dictionary.pkl', 'rb') as fichier:
        dico = pickle.load(fichier)
except :
    dico = {}


#%% Creation of a class for the user interface
class Application(tk.Frame):
    #%% Utility functions
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.create_widgets()
        self.master.title("Log in")


    def create_widgets(self):
        
        self.username_frame = tk.Frame(self.master)
        self.username_frame.grid(row=0, column=0, padx=10, pady=(10,0), sticky="w")
        self.username_label = tk.Label(self.username_frame, text="Username :")
        self.username_label.pack(side="left")
        self.username_entry = tk.Entry(self.username_frame, width = 35)
        self.username_entry.pack(side="left")
        self.username_entry.insert(0, "ado")
        
        self.password_frame = tk.Frame(self.master)
        self.password_frame.grid(row=0, column=0, padx=(300,10), pady=(10,0), sticky="w")
        self.password_label = tk.Label(self.password_frame, text="Password :")
        self.password_label.pack(side="left")
        self.password_entry = tk.Entry(self.password_frame, show="*",width = 35)
        self.password_entry.pack(side="left")
        self.password_entry.insert(0, "BOC4future++")
        
        self.baseurl_frame = tk.Frame(self.master)
        self.baseurl_frame.grid(row=1, column=0, padx=15, pady=(10,0), sticky="w")
        self.baseurl_label = tk.Label(self.baseurl_frame, text="Base URL (e.g. : http://localhost:9584/ADONIS15_0 ) :")
        self.baseurl_label.pack(side="left")
        self.baseurl_entry = tk.Entry(self.baseurl_frame, width = 60)
        self.baseurl_entry.pack(side="left")
        self.baseurl_entry.insert(0, "http://localhost:9584/ADONIS15_0")
        
        self.continue_button = tk.Button(self.master, text="Analyze data", command=self.run_script_and_open_new_window, width = 30)
        self.continue_button.grid(row=2, columnspan=2, pady = (10,0), padx = (50,10))
        
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)
        self.master.grid_columnconfigure(2, weight=1)
        
        
    def open_new_window(self):
        self.new_window = tk.Toplevel(self.master)
        self.new_window.geometry("300x200")
        tk.Label(self.new_window, text="The script is running... Please wait \n Running script might take a while").pack()
        self.new_window.title("Analyzing data")
        
    def open_final_window(self):
        final_window = tk.Toplevel(self.master)
        final_window.geometry("300x200")
        tk.Label(final_window, text="Data analyzed").pack()
    
    def run_script_and_open_new_window(self):
        self.master.withdraw()
        self.open_new_window()
        self.master.update()
        self.run_script()
        self.new_window.destroy()
        self.master.update()
        sys.exit()

    
 
    #%% Scripts
    def run_script(self):
        #%%Global variables (To complete by the user)
        base_url = self.baseurl_entry.get()
        username = self.username_entry.get()
        password = self.password_entry.get()
        base_url = base_url + "/rest/2.0/"
        url = base_url + "repos/"
        response = requests.get(url, auth=(username,password))
        data = response.json()
        repos = data["repos"]
        ids = [repo["id"].replace("{", "").replace("}", "") for repo in repos]
        ListObjects = ['ACTOR', 'AGGREGATION', 'APPLICATION', 'ATTRIBUTE', 'BLOCK', 'CONNECTION', 'CONTROL', 'CONTROL_OBJECTIVE', 'CROSS_REFERENCE', 'DATA_OBJECT', 'DOCUMENT', 
                       'END_EVENT', 'ENTITY', 'EXCLUSIVE_GATEWAY', 'EXTERNAL_PARTNER_R', 'GROUP', 'INFRASTRUCTURE_ELEMENT', 'INITIATIVE', 'INTERFACE', 'INTERMEDIATE_EVENT', 
                       'INTERMEDIATE_EVENT_BOUNDARY', 'LABEL', 'LANE', 'LANE_VERTICAL', 'MESSAGE', 'NON_EXCLUSIVE_GATEWAY', 'NOTE', 'OPERATION', 'ORGANIZATIONAL_UNIT', 
                       'PERFORMANCE', 'PERFORMANCE_INDICATOR', 'PERFORMANCE_INDICATOR_OVERVIEW', 'PERFORMER', 'POOL', 'POOL_VERTICAL', 'PROCESS', 'PROCESS_START', 'PRODUCT', 
                       'PRODUCT_COMPONENT', 'RESOURCE_R', 'RISK', 'RISK_GROUP', 'ROLE', 'SERVICE', 'START_EVENT', 'SUB_CHOREOGRPHY', 'SUB_CONVERSATION', 'SUBPROCESS', 
                       'SUB_PROCESS', 'SWIMLANE_HORIZONTAL', 'SWIMLANE_VERTICAL', 'SYSTEM_BOUNDARY', 'TASK', 'USE_CASE', 'USER']
        ListObjectsHistory = ['ACTOR','APPLICATION','ATTRIBUTE','CONTROL','CONTROLE_OBJECTIVE','DOCUMENT','ENTITY','EXTERNAL_PARTNER_R',
                              'INFRASTRUCTURE_ELEMENT','INITIATIVE','INTERFACE','ORGANIZATIONAL_UNIT','PERFORMANCE_INDICATOR','PERFORMANCE_INDICATOR',
                              'PERFORMER','PROCESS','PRODUCT_COMPONENT','RESOURCE_R','RISK','RISK_GROUP','ROLE','SERVICE','USE_CASE']
        ListRepositoryID = ids
        #%% Main functions
        
    
        def Calculate_Percentage_One_Object_One_Repository(Object, RepoID):
            #Request the data
            repo_id = RepoID
            if Object != 'USER':
                endpoint = f'repos/{repo_id}/search?range-start=0&query=%7Bfilters%3A%5B%7B%22className%22%3A%22C_' + Object +'%22%7D%5D%7D&range-end=-1&'
            else :
                endpoint = f'repos/{repo_id}/search?range-start=0&query=%7Bfilters%3A%5B%7B%22className%22%3A%22' + Object +'%22%7D%5D%7D&range-end=-1&'
            response = requests.get(base_url + endpoint, auth=(username, password))
            if response.status_code == 200:
                global NameToAdd
                data = response.json()
                items = data['items']
                # Variable initialization and export folder creation
                attribute_counter = Counter() 
                NbObjects = len(items)
                Default_Values = ["Not specified", "", "No entry", "None", "No assessment data available to calculate as-is average", 0, '1970-01-01T01:00:00']
                ThisPath = os.path.dirname(os.path.abspath(__file__))
                NameToAdd = base_url.rsplit('/', 5)[1].replace(':', '_')
                Export_Folder_Path = ThisPath + '\\Results_' + NameToAdd + '\\Data_By_Object'
                if not os.path.exists(Export_Folder_Path):
                    os.makedirs(Export_Folder_Path)
                File_Mean = Export_Folder_Path + '\\Mean_Data_' + Object + '.xlsx'
                #For each object
                if len(items)>0 :
                    for item in items:
                        attributes = item['attributes']
                        alreadychanged = ['-']
                        idxchangeHistory = FindChangeHistoryInList(attributes)
                        for attribute in attributes:
                            idxchangeHistory = FindChangeHistoryInList(attributes)
                            if idxchangeHistory>-1 :
                                History = attributes[idxchangeHistory]['value']
                                for change in History :
                                    change = change['cells']
                                    attributechanged = change[FindIndexAttributeInList(change)]['value'].upper()
                                    if attributechanged in list(dico.keys()):
                                        attributechanged = dico[attributechanged]
                                    else :
                                        dico[attributechanged] = translate(attributechanged,'en').upper()
                                        attributechanged = dico[attributechanged]
                                        with open('dictionary.pkl', 'wb') as fichier:
                                            pickle.dump(dico, fichier)    
                                    if not attributechanged in alreadychanged:
                                        attribute_counter[attributechanged] += 1
                                        alreadychanged.append(attributechanged)
                                for attribute in attributes:
                                    # We want the name of the attribute
                                    attribute_name = attribute['name'].upper()
                                    if not attribute_name in alreadychanged:
                                        attribute_counter[attribute_name] += 0
                                        
                            else : 
                                # We want the name of the attribute
                                attribute_name = attribute['name'].upper()
                                if attribute_name in list(dico.keys()):
                                    attribute_name = dico[attribute_name]
                                else :
                                    dico[attribute_name] = translate(attribute_name,'en').upper()
                                    attribute_name = dico[attribute_name]
                                    with open('dictionary.pkl', 'wb') as fichier:
                                        pickle.dump(dico, fichier)    
                                if 'value' in attribute.keys():
                                    attribute_value = attribute['value']
                                    if not attribute_value in Default_Values :
                                        attribute_counter[attribute_name] += 1
                                    else :
                                        attribute_counter[attribute_name] += 0
                                elif 'targets' in attribute.keys():
                                    attribute_targets = attribute['targets'][0]
                                    if attribute_targets['name'] != "" :
                                        attribute_counter[attribute_name] += 1
                                    else :
                                        attribute_counter[attribute_name] += 0
                                else:
                                    attribute_counter[attribute_name] += 0

                    count_df = pd.DataFrame({'Attributes Name': list(attribute_counter.keys()), 'Count': list(attribute_counter.values())})
                    count_df['Proportion'] = count_df['Count']*100/NbObjects
                    count_df_sorted = count_df.sort_values('Proportion', ascending=False)
                    # count_df_sorted is a DataFrame with 3 columns : Attribute Name, Count (Number of times the attribute has been used), Proportion 
                    NbRows = count_df_sorted.shape[0]
               
                    if not os.path.exists(File_Mean):
                        CoefList = [NbObjects]*NbRows
                        Mean = pd.DataFrame({ 'Attributes Name' :  count_df_sorted['Attributes Name'], 'Percentage of use' : count_df_sorted['Proportion'], 'Coef' : CoefList})
                        Mean.to_excel(File_Mean)
                        Next_Coef = NbObjects
                    else :
                        Mean = pd.read_excel(File_Mean)
                        if Mean.shape[0] > 0 : 
                            Next_Coef = Mean['Coef'].values.max() + NbObjects
                        else :
                            Next_Coef = 0
                        for index, row in count_df_sorted.iterrows():
                            attribute_name = row['Attributes Name']
                            count = row['Count']
                            if attribute_name in Mean['Attributes Name'].values:
                                mean_index = Mean.index[Mean['Attributes Name'] == attribute_name][0]
                                mean_percentage_of_use = Mean.loc[mean_index, 'Percentage of use']
                                Mean.at[mean_index, 'Percentage of use'] = (((Next_Coef-NbObjects) * (mean_percentage_of_use/100) + count) / (Next_Coef))*100
                        Mean['Coef'] = Next_Coef
                        Mean.to_excel(File_Mean, index=False)

            # Error
            else:
                print('Error requesting data.')
                print(f'Response HTTP code: {response.status_code}')

        def Calculate_Mean_Percentage_Of_Use(List_Of_Repository_ID ,List_of_Objects):
            for obj in List_of_Objects:
                for repo in List_Of_Repository_ID :
                    user_input = obj
                    user_input2 = normalize(str(repo))
                    Calculate_Percentage_One_Object_One_Repository(user_input,user_input2)
                    print(f"Success for {obj} in the repository {repo}.")


        #%%Utility function
        def normalize(text):
            if not text.endswith("}"):
                text = text + '}'
            if text[0] != "{":
                text = '{' + text
            return text
        
        def FindChangeHistoryInList(L):
            idx = -1
            for i in range(len(L)):
                dic = L[i]
                if dic['name'] == 'Change history':
                    idx = i
                    break
            return(idx)

        def FindIndexAttributeInList(L):
            idx = -1
            for i in range(len(L)):
                dic = L[i]
                if dic['name'] == 'Attribute':
                    idx = i
                    break
            return(idx)
        
        def FindIndexLanguageInList(L):
            idx = -1
            for i in range(len(L)):
                dic = L[i]
                if dic['name'] == 'Language':
                    idx = i
                    break
            return(idx)
        
        def get_repos_ids():
            url = base_url + "repos/"
            response = requests.get(url, auth=(username,password))
            data = response.json()
            repos = data["repos"]
            ids = [repo["id"].replace("{", "").replace("}", "") for repo in repos]
            return ids

        #%%Graphical results in a PDF report
        def ShowResults():
            ThisPath = os.path.dirname(os.path.abspath(__file__))
            Folder_Path = ThisPath + '\\Results_' + NameToAdd + '\\Data_By_Object'
            model_research = Folder_Path + '\*.xlsx'
            File_List = glob.glob(model_research)
            pdf_pages = PdfPages(ThisPath + '\\Results_'+ NameToAdd +'\\PDFReport.pdf')
            for file in File_List :
                try :
                    count_df = pd.read_excel(file)
                    count_df_sorted = count_df.sort_values('Percentage of use', ascending=False)
                    #count_df_sorted = count_df_sorted.drop('Coef', axis=1)
                    count_df_sorted = count_df_sorted.iloc[:, 1:]
                    count_df_sorted.to_excel(file, index=False)
                    thresholds = {'green': (50, float('inf')),
                                  'orange': (10, 50),
                                  'red': (-float('inf'), 10)}
                    filtered_Names = []
                    filtered_Props = []
                    colors = []
                    attributes_100_used = []
                    attributes_never_used = []
                    for (Name,Prop) in zip(count_df_sorted['Attributes Name'],count_df_sorted['Percentage of use']):
                        if (Prop  > 5) and (Prop < 80):
                            filtered_Names.append(Name)
                            filtered_Props.append(Prop)
                            for color, (lower, upper) in thresholds.items():
                                if lower <= Prop <= upper:
                                    colors.append(color)
                                    break
                        else : 
                            if (not Name in attributes_100_used) and (Prop >= 80):
                                attributes_100_used.append(Name)
                            if (not Name in attributes_never_used) and (Prop < 5):
                                attributes_never_used.append(Name)
                                
                    title = os.path.basename(file)
                    first_index = title.find("_")
                    title = title[first_index+1 :]
                    title = title[title.find("_")+1 : title.find(".")]
                    if title[-2:]=='_R':
                        title = title[:-2]
                        title = title + " WITH CHANGE HISTORY"
                    if title in ListObjectsHistory:
                        title = title + " WITH CHANGE HISTORY"
                    title = title.replace("_"," ")
                    
                    
                    plt.figure(1,figsize=(30,25))
                    plt.subplot(2, 1,2)
                    plt.bar(filtered_Names, filtered_Props, color = colors)
                    plt.bar_label(plt.bar(filtered_Names, filtered_Props, color=colors), labels=[f"{val:.2f}" for val in filtered_Props])
                    plt.title(title + ' :\nMEAN OF USE FOR EVERY ATTRIBUTES')
                    plt.xticks(rotation=45, ha = 'right')
                    
                    plt.subplot(2,2,1)
                    if len(attributes_100_used) > 22 :
                        plt.text(0.5, 0.5, ''.join(attributes_100_used[i]+' | ' if ( ((i%3==0) or (i%3==1)) and (i < len(attributes_100_used)-2) and (len(attributes_100_used[i])+len(attributes_100_used[i+1])+len(attributes_100_used[i+2])<90)) else attributes_100_used[i] + '\n' for i in range(len(attributes_100_used))), verticalalignment='center', horizontalalignment='center', fontsize=13, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round'))
                    else :
                        plt.text(0.5, 0.5, '\n'.join(attributes_100_used), verticalalignment='center', horizontalalignment='center', fontsize=13, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round'))
                    plt.axis('off')
                    plt.title(title + ' :\nATTRIBUTES USED MORE THAN 80%', y=1.05)
                    
                    plt.subplot(2,2,2)
                    if len(attributes_never_used) > 22 :
                        plt.text(0.5, 0.5, ''.join(attributes_never_used[i]+' | ' if ( ((i%3==0) or (i%3==1)) and (i < len(attributes_never_used)-2) and (len(attributes_never_used[i])+len(attributes_never_used[i+1])+len(attributes_never_used[i+2])<90)) else attributes_never_used[i] + '\n' for i in range(len(attributes_never_used))), verticalalignment='center', horizontalalignment='center', fontsize=13, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round'))
                    else :
                        plt.text(0.5, 0.5, '\n'.join(attributes_never_used), verticalalignment='center', horizontalalignment='center', fontsize=13, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round'))
                    plt.axis('off')
                    plt.title(title + ' :\nATTRIBUTES USED LESS THAN 5%', y=1.05)
                    pdf_pages.savefig()
                    plt.close()
                    
                except Exception as e :
                    print(f"Problem with {os.path.basename(file)} : {e}")
                
            pdf_pages.close()
                    
        def WhichOnesUnder5():
            ThisPath = os.path.dirname(os.path.abspath(__file__))
            Folder_Path = ThisPath + '\\Results_' + NameToAdd + '\\Data_By_Object'
            model_research = Folder_Path + '\*.xlsx'
            File_List = glob.glob(model_research)
            Attributes_Under_5 = pd.DataFrame()
            for file in File_List :
                title = os.path.basename(file)
                title = title[title.find("_")+1 : title.find(".")]
                if title[-2:]=='_R':
                    title = title[:-2]
                    title = title + " WITH CHANGE HISTORY"
                if title in ListObjectsHistory:
                    title = title + " WITH CHANGE HISTORY"
                title = title.replace("_"," ")
                count_df = pd.read_excel(file)
                count_df_sorted = count_df.sort_values('Percentage of use', ascending=False)
                filtered_Names = []
                for (Name,Prop) in zip(count_df_sorted['Attributes Name'],count_df_sorted['Percentage of use']):
                    if Prop < 5:
                        filtered_Names.append(Name)
                filtered_Names = filtered_Names + [""]*200
                Attributes_Under_5[title]=pd.Series(filtered_Names)
                count_df_sorted.to_excel(file)
            ExportFolder = ThisPath + '\\Results_' + NameToAdd + '\\Attributes_Under_5%_of_Use.xlsx'
            Attributes_Under_5.to_excel(ExportFolder)
            
            return('Sucess Attributes under 5%')

        #%% Run
        Calculate_Mean_Percentage_Of_Use(ListRepositoryID, ListObjects)
        ShowResults()
        WhichOnesUnder5()
        
root = tk.Tk()
root.geometry("590x120")
app = Application(master=root)
app.mainloop()
root.quit()
