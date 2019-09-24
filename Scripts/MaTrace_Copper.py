# -*- coding: utf-8 -*-
"""
Created on Fri Mar  9 10:49:33 2018

@author: sklose
"""

import logging
import os
import numpy as np
import time
import datetime
import scipy.io
import scipy.stats   
import matplotlib.pyplot as plt    
import pandas as pd
import shutil 
import uuid
import xlrd


from scipy.optimize import curve_fit
from scipy.integrate import simps
from numpy import trapz
from openpyxl import load_workbook

#%% Define loggig and main Paths
def function_logger(file_level, Name_Scenario, Path_Result, console_level): # Set up logger
    # remove all handlers from logger
    logger = logging.getLogger()
    logger.handlers = [] # required if you don't want to exit the shell
    logger.setLevel(logging.DEBUG) #By default, logs all messages

    if console_level != None:
        console_log = logging.StreamHandler() #StreamHandler logs to console
        console_log.setLevel(console_level)
        console_log_format = logging.Formatter('%(message)s') # ('%(asctime)s - %(message)s')
        console_log.setFormatter(console_log_format)
        logger.addHandler(console_log)

    file_log = logging.FileHandler(Path_Result + '\\' + Name_Scenario + '.html', mode='w', encoding=None, delay=False)
    file_log.setLevel(file_level)
    #file_log_format = logging.Formatter('%(asctime)s - %(lineno)d - %(levelname)-8s - %(message)s<br>')
    file_log_format = logging.Formatter('%(message)s<br>')
    file_log.setFormatter(file_log_format)
    logger.addHandler(file_log)
    return logger, console_log, file_log
      
def ensure_dir(f): # Checks whether a given directory f exists, and creates it if not
    d = os.path.dirname(f)
    if not os.path.exists(d):
        os.makedirs(d) 
        
def greyfade(color_triple,greyfactor): # returns a color triple (RGB) that is between the original color triple and and equivalent grey value, linear scaling: greyfactor = 0: original, greyfactor = 1: grey
    return (1 - greyfactor) * color_triple + greyfactor * color_triple.sum() / 3


"""
Set main path
"""    

Path = '..\\'
    
Path_List=[]
Input_Paths = open('Path_names.txt','r')
for line in Input_Paths.read().split('\n'):
    Path_List.append(line)
Input_Paths.close()    
    
    
Project_MainPath = Path_List[1]
  
Name_User        = Path_List[0]

Input_Data        = Path_List[2]

Path_Data   = Project_MainPath + 'Data\\'
Path_Results = Project_MainPath + 'Results\\'
Path_Script = Project_MainPath + 'Scripts\\'


#%% Read Configuarion file    

Project_DataFileName = Input_Data
Project_DataFilePath = Path_Data + Project_DataFileName 
Project_DataFile_WB  = xlrd.open_workbook(Project_DataFilePath)
Project_Configsheet  = Project_DataFile_WB.sheet_by_name('Scenario_Overview')





                
Name_Scenario      = Project_Configsheet.cell_value(5,2)
Number_Scenario      = Project_Configsheet.cell_value(3,2)
StartTime          = datetime.datetime.now()
TimeString         = str(StartTime.year) + '_' + str(StartTime.month) + '_' + str(StartTime.day) + '__' + str(StartTime.hour) + '_' + str(StartTime.minute) + '_' + str(StartTime.second)
Path_Result        = Path_Results + Name_Scenario + '_' + TimeString + '\\'

# Read control and selection parameters into dictionary
ScriptConfig = {'Scenario_Description': Project_Configsheet.cell_value(6,2)}
ScriptConfig['Scenario Name'] = Name_Scenario
for m in range(9,35): # add all defined control parameters to dictionary
    try:
        ScriptConfig[Project_Configsheet.cell_value(m,1)] = np.int(Project_Configsheet.cell_value(m,2))
    except:
        ScriptConfig[Project_Configsheet.cell_value(m,1)] = str(Project_Configsheet.cell_value(m,2))
        
ScriptConfig['Current_UUID'] = str(uuid.uuid4())

# Create scenario folder
ensure_dir(Path_Result)
#Copy script and Config file into that folder
shutil.copy(Project_DataFilePath, Path_Result + Project_DataFileName)
shutil.copy(Project_MainPath + 'Scripts\\MaTrace_Copper.py', Path_Result + 'MaTrace_Copper.py')
# Initialize logger    
[Mylog,console_log,file_log] = function_logger(logging.DEBUG, Name_Scenario + '_' + TimeString, Path_Result, logging.DEBUG) 

# log header and general information
Mylog.info('<html>\n<head>\n</head>\n<body bgcolor="#ffffff">\n<br>')
Mylog.info('<font "size=+5"><center><b>Script ' + 'MaTrace_Copper' + '.py</b></center></font>')
Mylog.info('<font "size=+5"><center><b>Version: 2018-03-12 or later.</b></center></font>')
Mylog.info('<font "size=+4"> <b>Current User: ' + Name_User + '.</b></font><br>')
Mylog.info('<font "size=+4"> <b>Current Path: ' + Project_MainPath + '.</b></font><br>')
Mylog.info('<font "size=+4"> <b>Current Scenario: ' + Name_Scenario + '.</b></font><br>')
Mylog.info(ScriptConfig['Scenario_Description'])
Mylog.info('Unique ID of scenario run: <b>' + ScriptConfig['Current_UUID'] + '</b>')

Time_Start = time.time()
Mylog.info('<font "size=+4"> <b>Start of simulation: ' + time.asctime() + '.</b></font><br>')


#%% Read model parameters        

Mylog.info('<p><b>Reading model definitions.</b></p>')
Project_DefSheet  = Project_DataFile_WB.sheet_by_name('Definitions')

Par_NoOfProducts  = int(Project_DefSheet.cell_value(2,5))
Mylog.info('<p>Number of products: ' + str(Par_NoOfProducts) + '.</p>')
Par_NoOfProductGroups  = int(Project_DefSheet.cell_value(2,3))
#Mylog.info('<p>Number of product groups: ' + str(Par_No_int(Project_DefSheet.cell_value(2,2)) +ProductGroups) + '.</p>')
Par_NoOfYears     = int(Project_DefSheet.cell_value(2,2)) #[2015;2100]
Mylog.info('<p>Number of years: ' + str(Par_NoOfYears) + '.</p>')
Par_NoOfScraps    = int(Project_DefSheet.cell_value(2,7)) #[0: new scrap, 1: old scrap]
Mylog.info('<p>Number of scrap types: ' + str(Par_NoOfScraps) + '.</p>')
Par_NoOfRecyclingRoutes   = int(Project_DefSheet.cell_value(2,6)) #[0: new scrap, 1: old scrap]
Mylog.info('<p>Number of Recycling Routes: ' + str(Par_NoOfRecyclingRoutes) + '.</p>')
Par_NoOfSecMetals = int(Project_DefSheet.cell_value(2,8)) #[0: BOF route, 1: EAF route]
Mylog.info('<p>Number of refinement processes: ' + str(Par_NoOfSecMetals) + '.</p>')
Par_NoOfRegions   = int(Project_DefSheet.cell_value(2,9)) #[0: BOF route, 1: EAF route]
if Par_NoOfRegions==1:
    Mylog.info('Number of Regions: global ' + '.</p>')
else: Mylog.info('<p>Number of Regions: ' + str(Par_NoOfRegions) + '.</p>')


Def_ProductNames      = []
for m in range (0,Par_NoOfProducts):
    Def_ProductNames.append(Project_DefSheet.cell_value(m+4,3))    

Par_Time = [] # Time vector of the model, unit: year, first element: 2015
for m in range(0,Par_NoOfYears):
    Par_Time.append(np.int(Project_DefSheet.cell_value(m+4,2)))
    
if Par_NoOfRegions==1:
    Project_Datasheet  = Project_DataFile_WB.sheet_by_name('Parameters_global')
    
else:
    Project_Datasheet  = Project_DataFile_WB.sheet_by_name('Parameters_regions')
   
Mylog.info('Read and format parameters for manufacture and use phase.<br>')

#%% Read Efficiency Parameters

if ScriptConfig['Modus'] == 'Trace 2015 mined copper': 
    Global_Copper_Production2015 = Project_Datasheet.cell_value(5,2)  
   
Copper_use_in_regions_share=np.zeros(Par_NoOfRegions)
for r in range(0,Par_NoOfRegions):
    Copper_use_in_regions_share[r]=Project_Datasheet.cell_value(5+(Par_NoOfProducts+1)*r,3)
    
Copper_use_in_regions=np.zeros(Par_NoOfRegions)
for r in range(0,Par_NoOfRegions):
    Copper_use_in_regions[r]=Project_Datasheet.cell_value(5+(Par_NoOfProducts+1)*r,3)*Global_Copper_Production2015
    
Par_D_AllocationCopperToProducts = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts,Par_NoOfSecMetals)) # This is the D-matrix
for m in range(0,Par_NoOfYears):
    for r in range(0,Par_NoOfRegions):
        for p in range(0,Par_NoOfProducts):
            for c in range(0,Par_NoOfSecMetals):
                Par_D_AllocationCopperToProducts[m,r,p,c]  = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*r,7+c)

    
Par_Input_F_0_8 = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))  #Fabrication yield in regions
Par_Input_F_0_8[0,:,:] = np.einsum('ij,i->ij',Par_D_AllocationCopperToProducts[0,:,:,0],Copper_use_in_regions[:]) 

for m in range (1,Par_NoOfYears):
    for r in range(0,Par_NoOfRegions):
        for p in range(0,Par_NoOfProducts):
            Par_Input_F_0_8[m,:,:] = 0   
            
            
Par_Lambda_Fabrication_Efficiency = np.zeros((Par_NoOfProducts,Par_NoOfRegions))    # fabrication yield for products
for p in range(0,Par_NoOfProducts):
   for r in range(0,Par_NoOfRegions):
       Par_Lambda_Fabrication_Efficiency[p,r] = Project_Datasheet.cell_value(5+p+(Par_NoOfProducts+1)*r,12)

if ScriptConfig['Sensitivity'] == 'Sens fabrication eff long':
    Mylog.info('Sens fabrication eff  short.<br>')
    for p in range(0,Par_NoOfProducts):
        for r in range(0,Par_NoOfRegions): 
                Par_Lambda_Fabrication_Efficiency[p,r] = Par_Lambda_Fabrication_Efficiency[p,r] * 1.1
                if Par_Lambda_Fabrication_Efficiency[p,r] > 1:
                    Par_Lambda_Fabrication_Efficiency[p,r] = 1

if ScriptConfig['Sensitivity'] == 'Sens fabrication eff short':
    Mylog.info('Sens fabrication eff  short.<br>')
    for p in range(0,Par_NoOfProducts):
        for r in range(0,Par_NoOfRegions): 
                Par_Lambda_Fabrication_Efficiency[p,r] = Par_Lambda_Fabrication_Efficiency[p,r] * 0.9


Par_Tau = np.zeros((Par_NoOfRegions,Par_NoOfProducts)) # Mean lifetime array, unit: year
for r in range(0,Par_NoOfRegions):
    for p in range(0,Par_NoOfProducts):
        Par_Tau[r,p] = Project_Datasheet.cell_value(5+p+(Par_NoOfProducts+1)*r,13)

#increase lifetime if lifetime flag is set:
if ScriptConfig['Lifetime extension'] == 'all products':
    Par_Tau = 1.2 * Par_Tau
    Mylog.info('Lifetime extention 20%.<br>')
#read standart deviaion of lifetime
Par_StandDev_Tau=float(Project_Datasheet.cell_value(2,13))

if ScriptConfig['Lifetime extension'] == 'all products in IC':
    for p in range(0,Par_NoOfProducts):
        Par_Tau[0,p] = 1.2 * Par_Tau[0,p]    
        Par_Tau[4,p] = 1.2 * Par_Tau[4,p] 
        Par_Tau[5,p] = 1.2 * Par_Tau[5,p] 
        
if ScriptConfig['Lifetime extension'] == 'C&E':
    for r in range(0,Par_NoOfRegions):
        for p in range(12,17):
            Par_Tau[r,p] = 1.2 * Par_Tau[r,p]  
        
if ScriptConfig['Lifetime extension'] == 'C&E in IC':
   for p in range(12,17):
           Par_Tau[0,p] = 1.2 * Par_Tau[0,p]    
           Par_Tau[4,p] = 1.2 * Par_Tau[4,p] 
           Par_Tau[5,p] = 1.2 * Par_Tau[5,p] 
           
if ScriptConfig['Lifetime extension'] == 'Decreased C&E':
  for r in range(0,Par_NoOfRegions):
    for p in range(12,17):
           Par_Tau[r,p] = 0.8 * Par_Tau[r,p]    
          
if ScriptConfig['Lifetime extension'] == 'Moderate':
  for r in range(0,Par_NoOfRegions):
    for p in range(12,17):
           Par_Tau[r,p] = 2 * Par_Tau[r,p]    
           
if ScriptConfig['Lifetime extension'] == 'Ambitious':
  for r in range(0,Par_NoOfRegions):
    for p in range(12,17):
           Par_Tau[r,p] = 3.5 * Par_Tau[r,p]    

if ScriptConfig['Sensitivity'] == 'Sens Lifetime long':
    Par_Tau = 1.1 * Par_Tau
    Mylog.info('sens lifetime long.<br>')

if ScriptConfig['Sensitivity'] == 'Sens Lifetime short':
    Par_Tau = 0.9 * Par_Tau
    Mylog.info('Sens lifetime short.<br>')

Par_Sigma = Par_StandDev_Tau * Par_Tau # Standard deviation of lifetime array, region by product, unit: year       

Mylog.info('Define lifetime distribution of product cohorts in MaTrace_pdf Matrix.<br>')
MaTrace_pdf = np.zeros((Par_NoOfYears,Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts)) 
AgeMatrix   = np.zeros((Par_NoOfYears,Par_NoOfYears))
for c in range(0, Par_NoOfYears):  # cohort index
    for y in range(c + 1, Par_NoOfYears):
        AgeMatrix[y,c] = y-c
for R in range(0,Par_NoOfRegions):
    for P in range(0,Par_NoOfProducts):
        MaTrace_pdf[:,:,R,P] = scipy.stats.norm(Par_Tau[R,P], Par_Sigma[R,P]).pdf(AgeMatrix)  # Call scipy's Norm function with Mean, StdDev, and Age
# No discard in historic years and year 0:
for m in range(0,Par_NoOfYears):
    MaTrace_pdf[0:m+1,m,:,:] = 0

Mylog.info('Read parameters for obsolete products and waste management industries.<br>')
Par_Omega_ObsoleteStocks = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts)) # Par: Ω
for m in range(0,Par_NoOfYears):
    for r in range(0,Par_NoOfRegions):
        for p in range(0,Par_NoOfProducts):
            Par_Omega_ObsoleteStocks[m,r,p] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*r,14)
   
if ScriptConfig['Sensitivity'] == 'Sens omega high':
    Mylog.info('Sens omega high.<br>')
    for m in range(0,Par_NoOfYears):
        for r in range(0,Par_NoOfRegions):
            for p in range(0,Par_NoOfProducts):
                Par_Omega_ObsoleteStocks[m,r,p] = Par_Omega_ObsoleteStocks[m,r,p] * 1.1
            
if ScriptConfig['Sensitivity'] == 'Sens omega low':
    Mylog.info('Sens omega low.<br>')
    for m in range(0,Par_NoOfYears):
        for r in range(0,Par_NoOfRegions):
            for p in range(0,Par_NoOfProducts):
                Par_Omega_ObsoleteStocks[m,r,p] = Par_Omega_ObsoleteStocks[m,r,p] * 0.9

Par_Sigma_Losses = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts)) # Par: Ω
for m in range(0,Par_NoOfYears):
    for r in range(0,Par_NoOfRegions):
        for p in range(0,Par_NoOfProducts):
            Par_Sigma_Losses[m,r,p] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*r,15)

if ScriptConfig['Sensitivity'] == 'Sens sigma high':
    Mylog.info('Sens sigma high.<br>')
    for m in range(0,Par_NoOfYears):
        for r in range(0,Par_NoOfRegions):
            for p in range(0,Par_NoOfProducts):
                Par_Sigma_Losses[m,r,p] = Par_Sigma_Losses[m,r,p] * 1.1


if ScriptConfig['Sensitivity'] == 'sens sigma low':
    Mylog.info('sens sigma low.<br>')
    for m in range(0,Par_NoOfYears):
        for r in range(0,Par_NoOfRegions):
            for p in range(0,Par_NoOfProducts):
                Par_Sigma_Losses[m,r,p] = Par_Sigma_Losses[m,r,p] * 0.9

Par_Gamma_EoL_Collection_Rate_Copper = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))      
for m in range(0,Par_NoOfYears):
     for r in range(0,Par_NoOfRegions):
          for p in range(0,Par_NoOfProducts):
                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*r,16)


                    
#%% Read Process Transformation Parameters

# Allocation of EoL Products to scrap groups in different regions 
Par_A_EolToScrap_Copper = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts,Par_NoOfScraps))        
for m in range(0,Par_NoOfYears):
   for p in range(0,Par_NoOfProducts):
      for r in range(0,Par_NoOfRegions):
          for s in range(0,Par_NoOfScraps):
             Par_A_EolToScrap_Copper[m,r,p,s] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*r,18+s)
                
if ScriptConfig['WEEE consumer sorting'] == 'Improved':
    for m in range(0,Par_NoOfYears):
        for p in range(0,Par_NoOfProducts):
            for r in range(0,Par_NoOfRegions):
                for s in range(0,Par_NoOfScraps):
                    Par_A_EolToScrap_Copper[m,r,p,s] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*Par_NoOfRegions,18+s)
                
if ScriptConfig['WEEE consumer sorting'] == 'Improved in IC':
    for m in range(0,Par_NoOfYears):
        for p in range(0,Par_NoOfProducts):
                for s in range(0,Par_NoOfScraps):
                    Par_A_EolToScrap_Copper[m,0,p,s] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*Par_NoOfRegions,18+s)
                    Par_A_EolToScrap_Copper[m,4,p,s] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*Par_NoOfRegions,18+s)
                    Par_A_EolToScrap_Copper[m,5,p,s] = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*Par_NoOfRegions,18+s)
                  
Mylog.info('Read parameters for scrap treatment and refining.<br>')  

Par_Phi_Scrap_Sorting_Efficiency = np.zeros((Par_NoOfRegions,Par_NoOfScraps))  
for r in range(0,Par_NoOfRegions):
    for s in range(0,Par_NoOfScraps):
        Par_Phi_Scrap_Sorting_Efficiency[r,s] = Project_Datasheet.cell_value(s+5+(Par_NoOfProducts+1)*r,27)


                    
                    
if ScriptConfig['Increased sorting rate'] == 'Moderate':
    Mylog.info('Sens scrap sorting efficiency high.<br>')
    for r in range(0,Par_NoOfRegions):
        for s in range(0,Par_NoOfScraps):
                Par_Phi_Scrap_Sorting_Efficiency[r,s] = Par_Phi_Scrap_Sorting_Efficiency[r,s] * 1.1
                if Par_Phi_Scrap_Sorting_Efficiency[r,s] > 1:
                    Par_Phi_Scrap_Sorting_Efficiency[r,s]=1             
                    
                    
if ScriptConfig['Increased sorting rate'] == 'Ambitious':
    Mylog.info('Sens scrap sorting efficiency high.<br>')
    for r in range(0,Par_NoOfRegions):
        for s in range(0,Par_NoOfScraps):
                Par_Phi_Scrap_Sorting_Efficiency[r,s] = Par_Phi_Scrap_Sorting_Efficiency[r,s] * 1.3
                if Par_Phi_Scrap_Sorting_Efficiency[r,s] > 1:
                    Par_Phi_Scrap_Sorting_Efficiency[r,s]=1 
               
if ScriptConfig['Sensitivity'] == 'Sens scrap sorting efficiency low':
    Mylog.info('Sens scrap sorting efficiency low.<br>')
    for r in range(0,Par_NoOfRegions):
        for s in range(0,Par_NoOfScraps):
                Par_Phi_Scrap_Sorting_Efficiency[r,s] = Par_Phi_Scrap_Sorting_Efficiency[r,s] *  0.9


if ScriptConfig['Sensitivity'] == 'Sens scrap sorting efficiency high':
    Mylog.info('Sens scrap sorting efficiency high.<br>')
    for r in range(0,Par_NoOfRegions):
        for s in range(0,Par_NoOfScraps):
                Par_Phi_Scrap_Sorting_Efficiency[r,s] = Par_Phi_Scrap_Sorting_Efficiency[r,s] * 1.1
                if Par_Phi_Scrap_Sorting_Efficiency[r,s] > 1:
                    Par_Phi_Scrap_Sorting_Efficiency[r,s]=1

Par_Theta_Copper_recovery_from_scrap_in_recyclingroute = np.zeros((Par_NoOfRegions,Par_NoOfScraps,Par_NoOfRecyclingRoutes))
for r in range(0,Par_NoOfRegions):
   for s in range(0,Par_NoOfScraps):   
       for t in range (0,Par_NoOfRecyclingRoutes):
          Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[r,s,t] = Project_Datasheet.cell_value(s+5+(Par_NoOfProducts+1)*r,34+t)        
 

if ScriptConfig['Sensitivity'] == 'Sens recovery high':
    Mylog.info('Sens copper recovery high.<br>')
    for r in range(0,Par_NoOfRegions):
        for s in range(0,Par_NoOfScraps):   
            for t in range (0,Par_NoOfRecyclingRoutes):
                Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[r,s,t] = Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[r,s,t] * 1.1
                if Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[r,s,t] > 1:
                    Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[r,s,t] = 1
                    
                    
if ScriptConfig['Sensitivity'] == 'Sens recovery low':
    Mylog.info('Sens recovery low.<br>')
    for r in range(0,Par_NoOfRegions):
        for s in range(0,Par_NoOfScraps):   
            for t in range (0,Par_NoOfRecyclingRoutes):
                Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[r,s,t] = Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[r,s,t] * 0.9

        
Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf= np.zeros((Par_NoOfRegions))
for r in range(0,Par_NoOfRegions):
          Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[r] = Project_Datasheet.cell_value(5+(Par_NoOfProducts+1)*r,54)        


if ScriptConfig['IT efficiency'] == 'Moderate':
    Mylog.info('IT efficiency Moderate.<br>')
    for r in range(0,Par_NoOfRegions):
            Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[r]=0.8
          
                    
if ScriptConfig['IT efficiency'] == 'Ambitious':
    Mylog.info('IT efficiency Ambitious.<br>')
    for r in range(0,Par_NoOfRegions):
            Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[r]=0.95
    
Par_B_ScrapToRecyclingRoute = np.zeros((Par_NoOfRegions,Par_NoOfScraps,Par_NoOfRecyclingRoutes)) # Par: B2- scrap to recycling route 
for r in range(0,Par_NoOfRegions):
    for s in range(0,Par_NoOfScraps):
        for t in range(0,Par_NoOfRecyclingRoutes):
            Par_B_ScrapToRecyclingRoute[r,s,t] = Project_Datasheet.cell_value(s+5+(Par_NoOfProducts+1)*r,t+29) 
                    
Par_inf_collection_rate = np.zeros((Par_NoOfRegions,Par_NoOfProducts))  ##defines how much of EoL Products from Consumer and Electronic sector are collected informally
for r in range(0,Par_NoOfRegions):
   for p in range(0,Par_NoOfProducts):   
       Par_inf_collection_rate[r,p]  = Project_Datasheet.cell_value(p+5+(Par_NoOfProducts+1)*r,48)
       
Par_Phi_inf_Scrap_Sorting_Efficiency = np.zeros((Par_NoOfRegions))  ##defines the manual dissasembly efficiecy of the informal recycling sector
for r in range(0,Par_NoOfRegions):  
       Par_Phi_inf_Scrap_Sorting_Efficiency[r]  = Project_Datasheet.cell_value(5+(Par_NoOfProducts+1)*r,52)
      
if ScriptConfig['Sensitivity'] == 'Sens informal copper recovery high':
    Mylog.info('Sens informal copper recovery high.<br>')
    for r in range(0,Par_NoOfRegions): 
                Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[r]=Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[r] * 1.3

if ScriptConfig['Sensitivity'] == 'Sens informal copper recovery low':
    Mylog.info('Sens incineration rate low.<br>')
    for r in range(0,Par_NoOfRegions): 
                Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[r]=Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[r] * 0.7


if ScriptConfig['Sensitivity'] == 'Sens informal collection rate high':
    Mylog.info('Sens informal collection rate high.<br>')
    for r in range(0,Par_NoOfRegions): 
        for p in range(0, Par_NoOfProducts):
                Par_inf_collection_rate[r,p]=Par_inf_collection_rate[r,p] * 1.3
                if Par_Gamma_EoL_Collection_Rate_Copper[0,r,p] + Par_inf_collection_rate[r,p] > 1:
                    Par_inf_collection_rate[r,p]=1-Par_Gamma_EoL_Collection_Rate_Copper[0,r,p]
                    
                if Par_Gamma_EoL_Collection_Rate_Copper[0,r,p] + Par_inf_collection_rate[r,p] > 1:
                    Par_inf_collection_rate[r,p]=1-Par_Gamma_EoL_Collection_Rate_Copper[0,r,p]

if ScriptConfig['Sensitivity'] == 'Sens informal collection rate low':
    Mylog.info('Sens informal collection rate low.<br>')
    for r in range(0,Par_NoOfRegions): 
                Par_inf_collection_rate[r,:]=Par_inf_collection_rate[r,:] *  0.7
                
if ScriptConfig['Sensitivity'] == 'Sens Informal scrap sorting efficiency  low':
    Mylog.info('Informal scrap sorting efficiency low.<br>')
    for r in range(0,Par_NoOfRegions): 
                Par_Phi_inf_Scrap_Sorting_Efficiency[r]=Par_Phi_inf_Scrap_Sorting_Efficiency[r] * 0.7


if ScriptConfig['Sensitivity'] == 'Sens Informal scrap sorting efficiency  high':
    Mylog.info('Informal scrap sorting efficiency high.<br>')
    for r in range(0,Par_NoOfRegions): 
                Par_Phi_inf_Scrap_Sorting_Efficiency[r]=Par_Phi_inf_Scrap_Sorting_Efficiency[r] * 1.3
                if Par_Phi_inf_Scrap_Sorting_Efficiency[r] > 1:
                    Par_Phi_inf_Scrap_Sorting_Efficiency[r]=1
                
Par_C_RemeltingToSecondaryMetal = np.zeros((Par_NoOfRegions,Par_NoOfRecyclingRoutes,Par_NoOfSecMetals)) # Par: Theta2, Allocation of Remelted Scrap to Secondary Metals
for r in range(0,Par_NoOfRegions):
        for t in range(0,Par_NoOfRecyclingRoutes):
            for c in range(0,Par_NoOfSecMetals):
                Par_C_RemeltingToSecondaryMetal[r,t,c]  = Project_Datasheet.cell_value(t+5+(Par_NoOfProducts+1)*r,40+c)
 
if ScriptConfig['Sensitivity'] == 'Sens Collection rate high':
    Mylog.info('Sens collection rate high.<br>')
    for m in range(0,Par_NoOfYears):
        for r in range(0,Par_NoOfRegions):
            for p in range(0,12):
                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] = Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] * 1.1
            for p in range(12,16):
                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] = Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] * 1.2
              
            if Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] > 1:
                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p]=1
                
if ScriptConfig['Sensitivity'] == 'Sens Collection rate low':
    Mylog.info('Sens collection rate low.<br>')
    for m in range(0,Par_NoOfYears):
        for r in range(0,Par_NoOfRegions):
            for p in range(0,12):
                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] = Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] * 0.9
            for p in range(12,16):
                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] = Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] * 0.8
                             
 # Improved WEEE  recovery efficiency (Γ) if WEEE Collection rate is set to improve in all regions:               
if ScriptConfig['WEEE collection rate'] == 'Ambitious':       
    Mylog.info('WEEE collection rate Ambitious.<br>')
    for m in range(0,Par_NoOfYears):
        for p in range(12,16):
           for r in range(0,Par_NoOfRegions):
                 Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] =0.85
                 
                 
                 if Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] > 1:
                    Par_Gamma_EoL_Collection_Rate_Copper[m,r,p]=1
                 
                 if Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] + Par_inf_collection_rate[r,p] > 1:
                    Par_inf_collection_rate[r,p]=1-Par_Gamma_EoL_Collection_Rate_Copper[m,r,p]
       

 # Improved WEEE  recovery efficiency (Γ) if WEEE Collection rate is set to improved in IC:               
if ScriptConfig['WEEE collection rate'] == 'Moderate':       
    Mylog.info('WEEE collection rate Moderate.<br>')
    for m in range(0,Par_NoOfYears):
        for r in range(0,Par_NoOfRegions):
            for p in range(0,Par_NoOfProducts):
                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] = Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] * 1.1
                
                if Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] > 1:
                    Par_Gamma_EoL_Collection_Rate_Copper[m,r,p]=1
              
                if Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] + Par_inf_collection_rate[r,p] > 1:
                    Par_inf_collection_rate[r,p]=1-Par_Gamma_EoL_Collection_Rate_Copper[m,r,p]



#if ScriptConfig['Sensitivity'] == 'Sens informal collection rate high':
#    Mylog.info('Sens inf collection rate high.<br>')
#    for m in range(0,Par_NoOfYears):
#        for r in range(0,Par_NoOfRegions):
#            for p in range(0,Par_NoOfProducts):
#                Par_inf_collection_rate[m,r,p] = Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] * 1.3
#                if Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] > 1:
#                    Par_Gamma_EoL_Collection_Rate_Copper[m,r,p]=1
#                
#if ScriptConfig['Sensitivity'] == 'Sens Collection rate low':
#    Mylog.info('Sens collection rate low.<br>')
#    for m in range(0,Par_NoOfYears):
#        for r in range(0,Par_NoOfRegions):
#            for p in range(0,Par_NoOfProducts):
#                Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] = Par_Gamma_EoL_Collection_Rate_Copper[m,r,p] * 0.9
# 
#


#%% Read reuse and trade data
###Reuse von C&E (consumer goods: 12) und Vehicles (non-electrical vehicles: 10)  
          
Mylog.info('Read parameters for ELV trade and WEEE trade.<br>')  

Project_Tradesheet  = Project_DataFile_WB.sheet_by_name('ReuseandTrade')

Par_Chi_Reuse = np.zeros((Par_NoOfProducts,Par_NoOfRegions,Par_NoOfRegions)) 
for p in range(12,16):
   for r in range(0,Par_NoOfRegions):
       for R in range(0,Par_NoOfRegions):
           Par_Chi_Reuse[p,r,R] = Project_Tradesheet.cell_value(r+2,2+R)

Par_Chi_Reuse_inf = np.zeros((Par_NoOfProducts,Par_NoOfRegions,Par_NoOfRegions)) 
for p in range(12,16):
   for r in range(0,Par_NoOfRegions):
       for R in range(0,Par_NoOfRegions):
           Par_Chi_Reuse_inf[p,r,R] = Project_Tradesheet.cell_value(r+2,14+R)

if ScriptConfig['Increased reuse'] == 'Moderate':      
  for p in range(10,16):
       for r in range(0,Par_NoOfRegions):
          for R in range(0,Par_NoOfRegions):
              if r == R:
                   Par_Chi_Reuse_inf[p,r,R] = Par_Chi_Reuse_inf[p,r,R] + 0.1


if ScriptConfig['Increased reuse'] == 'Ambitious':      
  for p in range(10,16):
       for r in range(0,Par_NoOfRegions):
          for R in range(0,Par_NoOfRegions):
              if r == R:
                  Par_Chi_Reuse_inf[p,r,R] = Par_Chi_Reuse_inf[p,r,R]+0.5
                  
Par_Psi_EoL_Trade = np.zeros((Par_NoOfProducts,Par_NoOfRegions,Par_NoOfRegions)) 
if Par_NoOfRegions==1:
    for p in range(0,Par_NoOfProducts):
        for r in range(0,Par_NoOfRegions):
            for R in range(0,Par_NoOfRegions):
                Par_Psi_EoL_Trade[p,r,R] = 1   
else:         
    for p in range(0,Par_NoOfProducts):
        for r in range(0,Par_NoOfRegions):
            for R in range(0,Par_NoOfRegions):
                      if r==R:
                          Par_Psi_EoL_Trade[p,r,R] = 1
                      else:
                              Par_Psi_EoL_Trade[p,r,R] = 0

    for p in range(12,16):
        for r in range(0,Par_NoOfRegions):
            for R in range(0,Par_NoOfRegions):
                Par_Psi_EoL_Trade[p,r,R] = Project_Tradesheet.cell_value(r+13,2+R)



Par_Psi_EoL_Trade_inf = np.zeros((Par_NoOfProducts,Par_NoOfRegions,Par_NoOfRegions)) 
if Par_NoOfRegions==1:
    for p in range(0,Par_NoOfProducts):
        for r in range(0,Par_NoOfRegions):
            for R in range(0,Par_NoOfRegions):
                Par_Psi_EoL_Trade_inf[p,r,R] = 1   
else:         
    for p in range(0,Par_NoOfProducts):
        for r in range(0,Par_NoOfRegions):
            for R in range(0,Par_NoOfRegions):
                      if r==R:
                          Par_Psi_EoL_Trade_inf[p,r,R] = 1
                      else:
                              Par_Psi_EoL_Trade_inf[p,r,R] = 0


                

if ScriptConfig['No informal trade '] == 'no':   
    for p in range(12,16):
        for r in range(0,Par_NoOfRegions):
            for R in range(0,Par_NoOfRegions):
                Par_Psi_EoL_Trade_inf[p,r,R] = Project_Tradesheet.cell_value(r+13,14+R)
                
#%% Define System variables
      
Mylog.info('Define system variables.<br>')
#Define Process and market flows
# Define external input vector F_0_8(t,r,p):
F_0_8  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define final consumption vector F_8_1(t,r,p):
F_8_1  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define flow of obsolete steel F_y(t,r,p) (internal flow, not visible in system definition!):
F_y    = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define addition to obsolete stocks, F_1_1a(t,r,p):
F_1_2  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define amount of electronic products collected by informal recycling sector:
F_1_2inf  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define flow of domestic EoL products treatment, F_2_3(t,r,p):
F_2_3  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts)) 
F_2_3inf= np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts)) 
# Define flow of exports for re-use, to be inserted into stocks in other regions, F_2_8(t,r,p):
F_2_8  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
F_2_8inf= np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define flow of (old) scrap from EoL treatment to scrap market, F_3_4(t,[s,r]): # NOTE: For every year t, this variable is a column vector with outer index s and inner index r.
F_3_4  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfScraps)) 
F_3_4inf = np.zeros((Par_NoOfYears,Par_NoOfRegions)) 
#Define flow of total material for remelting, by remelting route a and region, F_4_5a(t,[m,r]):
F_4_5  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfRecyclingRoutes))
F_4_5inf= np.zeros((Par_NoOfYears,Par_NoOfRegions))
#Define flow of total remelted material, by producer and region, F_5_6(t,[m,r]):
F_5_6  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfSecMetals)) 
F_5_6inf = np.zeros((Par_NoOfYears,Par_NoOfRegions)) 
#Define flow of total secondary material consumed, F_6_7(t,[m,r]):
F_6_7  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfSecMetals)) 
F_6_7inf = np.zeros((Par_NoOfYears,Par_NoOfRegions)) 
# Define flow of steel in recycled products, by producing region, F_7_8(t,[r,p]):
F_7_8  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts)) 
F_7_8inf  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts)) 

# Define fabrication scrap, f_7_4(t,[s,r])
F_7_4  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfScraps)) 

#Define flows to the environment
F_1_Env_Omega   = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
F_1_Env_Sigma   = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
F_3_Env_Phi     = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfScraps))
F_3_Env_Phi_inf = np.zeros((Par_NoOfYears,Par_NoOfRegions))
F_5_Env_Theta   = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfRecyclingRoutes))
F_5_Env_Theta_inf= np.zeros((Par_NoOfYears,Par_NoOfRegions))
F_1_Gamma       = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))

# Define stocks
# In use stock
S_1  = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Obsolete stock
S_Env_Omega = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
    
S_Env_Sigma= np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))

S_Env_Phi = np.zeros((Par_NoOfYears,Par_NoOfRegions))

S_Env_Theta= np.zeros((Par_NoOfYears,Par_NoOfRegions))

S_Gamma= np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))

S_Brass= np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfSecMetals))
# Total amount of copper in the system
S_Tot = np.zeros(Par_NoOfYears)

# Define Balance of use phase
Bal_1                = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define Balance of EoL flows
Bal_ObsoleteProducts = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
# Define Balance of Waste management industries:
Bal_3                = np.zeros((Par_NoOfYears,Par_NoOfRegions))
# Define Balance of Waste management industries in the informal recycling sector:
Bal_3inf                = np.zeros((Par_NoOfYears,Par_NoOfRegions))

# Define Balance of the remelting industries
Bal_5                = np.zeros((Par_NoOfYears,1))
# Define Balance of the fabrication sectors:
Bal_7                = np.zeros((Par_NoOfYears,Par_NoOfRegions))

# Define market balances
#Define Balance of Final products market
Bal_8    = np.zeros((Par_NoOfYears,Par_NoOfRegions,Par_NoOfProducts))
#Define Balance of EoL products market
Bal_2    = np.zeros((Par_NoOfYears,Par_NoOfProducts))
#Define balance of scrap markets
Bal_4    = np.zeros((Par_NoOfYears,1))
#Define balance of material markets
Bal_6    = np.zeros((Par_NoOfYears,1))

#Define system-wide mass balance:
Bal_System           = np.zeros(Par_NoOfYears)


#Define total Losses
Total_Losses_regions = np.zeros((Par_NoOfRegions))  
Total_Losses_share   = np.zeros((Par_NoOfRegions))  
Total_Losses_Omega   = np.zeros((Par_NoOfRegions)) 
Total_Losses_Sigma   = np.zeros((Par_NoOfRegions)) 
Total_Losses_Phi     = np.zeros((Par_NoOfRegions)) 
Total_Losses_Theta   = np.zeros((Par_NoOfRegions)) 
Total_Losses_Gamma   = np.zeros((Par_NoOfRegions)) 

EoLRR=np.zeros((Par_NoOfYears,Par_NoOfRegions))     
EoLRR_global=np.zeros((Par_NoOfYears)) 

#%% Peform caclulations
Mylog.info('<p><b>Performing model calculations.</b></p>')
SY = ScriptConfig['StartYear no.'] -2015 # index for start year
EY = ScriptConfig['Time horizon']  -2015 # index for end year
Mylog.info('<p>Modus: Trace fate of copper </p>')
Mylog.info('<p>Tracing fate of Copper  ' + ' starting in year ' + str(Par_Time[SY]) + ', ending in year ' + str(Par_Time[EY]) + '.</p>')

F_0_8[SY,:,:] = Par_Input_F_0_8[SY,:,:]

# perform a year-by-year computation of the outflow, recycling, and inflow
Mylog.info('<p>MaTrace Copper was successfully initialized. Starting to calculate the future material distribution.</p>')
for CY in range(0,EY+1): # CY stands for 'current year'
        
        Mylog.info('<p>Performing calculations for year ' + str(Par_Time[CY]) + '.</p>')

        F_8_1[SY,:,:] = F_0_8[SY,:,:]
        
        Mylog.info('Step 1: Determine flow of copper in EoL products .<br>')
        for r in range(0,Par_NoOfRegions):
            for p in range(0,Par_NoOfProducts):
                # Use MaTrace_pdf to determine convolution of historic inflow with lifetime distribution
                F_y[CY,r,p] = (F_8_1[0:CY,r,p] * MaTrace_pdf[CY,0:CY,r,p]).sum()       

        Mylog.info('Step 2: Calculate obsolete stocks, trade of obsolete products, and material sent to the waste management industries.<br>')
        # Determine losses due to dissipation
        F_1_Env_Sigma[CY,:,:]=np.einsum('ij,ij->ij',F_y[CY,:,:],Par_Sigma_Losses[CY,:,:])
        # Determine obsolete stocks:
        F_1_Env_Omega[CY,:,:]=np.einsum('rp,rp->rp',F_y[CY,:,:]-F_1_Env_Sigma[CY,:,:],Par_Omega_ObsoleteStocks[CY,:,:])
        #Determine Flow of EoL products collected for recovery
      ##  F_1_Gamma[CY,:,:]=np.einsum('ij,ij->ij',(F_y[CY,:,:]-F_1_2inf[CY,:,:]-F_1_Env_Omega[CY,:,:]-F_1_Env_Sigma[CY,:,:]),(1-Par_Gamma_EoL_Collection_Rate_Copper[CY,:,:]))
        # Determine flow to EoL market:
      ##  F_1_2[CY,:,:]= F_y[CY,:,:]-F_1_2inf[CY,:,:] - F_1_Gamma[CY,:,:] - F_1_Env_Omega[CY,:,:] - F_1_Env_Sigma[CY,:,:]
        F_1_2inf[CY,:,:]=np.einsum('ij,ij->ij',Par_inf_collection_rate[:,:], F_y[CY,:,:]-F_1_Env_Omega[CY,:,:]-F_1_Env_Sigma[CY,:,:])
        F_1_2[CY,:,:]= np.einsum('ij,ij->ij',(F_y[CY,:,:]-F_1_Env_Omega[CY,:,:]-F_1_Env_Sigma[CY,:,:]),(Par_Gamma_EoL_Collection_Rate_Copper[CY,:,:]))
        F_1_Gamma[CY,:,:]=F_y[CY,:,:]-F_1_2inf[CY,:,:] - F_1_2[CY,:,:] - F_1_Env_Omega[CY,:,:] - F_1_Env_Sigma[CY,:,:]
#        # Determine export for re-use:
        F_2_8[CY,:,:]  = np.einsum('ij,jik->kj',F_1_2[CY,:,:],Par_Chi_Reuse[:,:,:])
        
        F_2_8inf[CY,:,:]  = np.einsum('ij,jik->kj',F_1_2inf[CY,:,:],Par_Chi_Reuse_inf[:,:,:])
#        # Determine flow to waste sorting including export for waste recovery:
        F_2_3[CY,:,:]  = np.einsum('ij,jik->kj', np.einsum('ij,ji->ij',F_1_2[CY,:,:],(1-Par_Chi_Reuse[:,:,:].sum(axis=2))), Par_Psi_EoL_Trade[:,:,:]) 

        F_2_3inf[CY,:,:]  = np.einsum('ij,jik->kj', np.einsum('ij,ji->ij',F_1_2inf[CY,:,:],(1-Par_Chi_Reuse_inf[:,:,:].sum(axis=2))), Par_Psi_EoL_Trade_inf[:,:,:]) 
        
        
        Bal_1[CY,:,:] =   F_1_2[CY,:,:] +F_1_Env_Sigma[CY,:,:]+F_1_Env_Omega[CY,:,:]+ F_1_Gamma[CY,:,:] - F_y[CY,:,:]+F_1_2inf[CY,:,:]

        Bal_2[CY,:] = F_1_2[CY,:,:].sum() - F_2_8[CY,:,:].sum() - F_2_3[CY,:,:].sum() + F_1_2inf[CY,:,:].sum() - F_2_8inf[CY,:,:].sum() - F_2_3inf[CY,:,:].sum()



        Mylog.info('Step 3: EoL material recovery and lossed to landfills.<br>')
        # Determine flow of old scrap from EoL treatment to scrap market, aggregate over products    
        F_3_4[CY,:,:] = np.einsum('...ij,...i', Par_A_EolToScrap_Copper[CY,:,:,:],F_2_3[CY,:,:])*Par_Phi_Scrap_Sorting_Efficiency[:,:]
        
        F_3_4inf[CY,:] = F_2_3inf[CY,:,:].sum(axis=1)* Par_Phi_inf_Scrap_Sorting_Efficiency[:]
        # Determine flow of recovery losses to landfills
        F_3_Env_Phi[CY,:,:] = (1 - Par_Phi_Scrap_Sorting_Efficiency[:,:])*np.einsum('...ij,...i', Par_A_EolToScrap_Copper[CY,:,:,:],F_2_3[CY,:,:])
        # Determine mass balance of scrap recovery process
        
        F_3_Env_Phi_inf[CY,:] = (1 - Par_Phi_inf_Scrap_Sorting_Efficiency[:])*F_2_3inf[CY,:,:].sum(axis=1)

        Bal_3[CY,:] = F_2_3[CY,:,:].sum(axis = 1) + F_2_3inf[CY,:,:].sum(axis=1) - F_3_4[CY,:,:].sum(axis = 1)- F_3_4inf[CY,:]  - F_3_Env_Phi[CY,:,:].sum(axis = 1) - F_3_Env_Phi_inf[CY,:]
        
        Bal_3inf[CY,:] = F_2_3inf[CY,:,:].sum(axis = 1) - F_3_4inf[CY,:] - F_3_Env_Phi_inf[CY,:] 

        Mylog.info('Step 4: Refining, manufacturing, and re-distribution into products.<br>')
        
        F_4_5[CY,:,:]=(np.einsum('rst,rst->rt',Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[:,:,:],np.einsum('ijk,ij->ijk',Par_B_ScrapToRecyclingRoute[:,:,:],(F_3_4[CY,:,:]+F_7_4[CY-1,:,:]))))
        F_4_5inf[CY,:]= Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[:]* (F_3_4inf[CY,:])       
        F_5_Env_Theta[CY,:,:]=np.einsum('rst,rst->rt',(1-Par_Theta_Copper_recovery_from_scrap_in_recyclingroute[:,:,:]),np.einsum('ijk,ij->ijk',Par_B_ScrapToRecyclingRoute[:,:,:],(F_3_4[CY,:,:]+F_7_4[CY-1,:,:])))
        F_5_Env_Theta_inf[CY,:]= (1-Par_Theta_Copper_recovery_from_scrap_in_recyclingroute_inf[:])* (F_3_4inf[CY,:])
        F_5_6[CY,:,:]= np.einsum('...i,...ij->...j',F_4_5[CY,:,:],Par_C_RemeltingToSecondaryMetal[:,:,:])
        F_5_6inf[CY,:]=F_4_5inf[CY,:]          
        # Determine net production of goods from recycled material:
        F_6_7[CY,:,:]  = F_5_6[CY,:,:]
        F_6_7inf[CY,:]=F_5_6inf[CY,:]
        F_7_8inf[CY,:,:]=np.einsum('rp,r->rp', Par_D_AllocationCopperToProducts[CY,:,:,0],(F_6_7inf[CY,:]))
        # De  termine net production of goods from recycled material:    
        F_7_8[CY,:,:]  = np.einsum('rp,pr->rp',np.einsum('rpc,rc->rp', Par_D_AllocationCopperToProducts[CY,:,:,:],(F_6_7[CY,:,:])),Par_Lambda_Fabrication_Efficiency[:,:])

        F_7_4[CY,:,6]=np.einsum('rp,pr->rp',np.einsum('rpc,rc->rp', Par_D_AllocationCopperToProducts[CY,:,:,:],(F_5_6[CY,:,:])),(1-Par_Lambda_Fabrication_Efficiency[:,:])).sum(axis=1)
     
        
        Bal_4[CY,:]  =  F_7_4[CY-1,:,:].sum() + F_3_4[CY,:,:].sum() + F_3_4inf[CY,:].sum() - F_4_5[CY,:,:].sum() -  F_4_5inf[CY,:].sum()- F_5_Env_Theta_inf[CY,:].sum() - F_5_Env_Theta[CY,:,:].sum()       
        Bal_5[CY,:]  = F_4_5[CY,:,:].sum() - F_5_6[CY,:,:].sum()
      
        Bal_6[CY,:]  = F_5_6[CY,:,:].sum() - F_6_7[CY,:,:].sum()

        Bal_7[CY,:]  = F_6_7[CY,:,:].sum(axis=1) - F_7_8[CY,:,:].sum(axis=1) - F_7_4[CY,:,:].sum(axis=1) + F_6_7inf[CY,:] - F_7_8inf[CY,:,:].sum(axis=1)
  
        
        
        
        Mylog.info('Step 5: Re-insert recycled and re-used goods into the stock.<br>')
        # Determine the total consumption of products as sum of external input x_in, re-used products from all regions, and recycled products from all regions
        F_8_1[CY,:,:] = (F_7_8[CY,:,:]+F_2_8[CY,:,:]+F_0_8[CY,:,:]+F_2_8inf[CY,:,:]+ F_7_8inf[CY,:,:])

        Bal_8[CY,:,:]   = F_8_1[CY,:,:] - F_0_8[CY,:,:] -  F_2_8[CY,:,:] - F_7_8[CY,:,:] - F_2_8inf[CY,:,:]

                    

        
        EoLRR[CY,:] = F_5_6[CY,:,:].sum(axis=1)/F_y[CY,:,:].sum(axis=1)
        EoLRR_global[CY] = F_5_6[CY,:,:].sum()/F_y[CY,:,:].sum()
        
         
        Mylog.info('Step 6: Determine stocks and overall system balance.<br>')
        if CY == 0:
            Mylog.info('Stock determination for year 0.<br>')
            S_1[0,:,:]                  = F_8_1[0,:,:]    -F_y[0,:,:]    
            S_Env_Omega[0,:,:]          = F_1_Env_Omega[0,:,:]
            S_Env_Sigma[0,:,:]          = F_1_Env_Sigma[0,:,:]
            S_Env_Phi[0,:]            = F_3_Env_Phi[0,:,:].sum(axis=1) + F_3_Env_Phi_inf[0,:]
            S_Env_Theta[0,:]          = F_5_Env_Theta[0,:,:].sum(axis=1) + F_5_Env_Theta_inf[0,:]
            S_Gamma[0,:,:]              = F_1_Gamma[0,:,:]
            S_Brass[0,:,:]              = F_6_7[0,:,:]
            
        else:
            Mylog.info('Stock determination.<br>')
            S_1[CY,:,:]                 = S_1[CY-1,:,:]         + F_8_1[CY,:,:] - F_y[CY,:,:]          
            S_Env_Omega[CY,:,:]         = S_Env_Omega[CY-1,:,:] + F_1_Env_Omega[CY,:,:]
            S_Env_Sigma[CY,:,:]         = S_Env_Sigma[CY-1,:,:] + F_1_Env_Sigma[CY,:,:]
            S_Env_Phi[CY,:]           = S_Env_Phi[CY-1,:]  + F_3_Env_Phi[CY,:,:].sum(axis=1) + F_3_Env_Phi_inf[CY,:]
            S_Env_Theta[CY,:]         = S_Env_Theta[CY-1,:] + F_5_Env_Theta[CY,:,:].sum(axis=1) + F_5_Env_Theta_inf[CY,:]  
            S_Gamma[CY,:,:]             = S_Gamma[CY-1,:,:]     + F_1_Gamma[CY,:,:]
            S_Brass[CY,:,:]             = S_Brass[CY-1,:,:]     + F_6_7[CY,:,:]
            
           # Bal_1[CY,:,:] = F_8_1[CY,:,:]  - F_y[CY,:,:] - (S_1[CY,:,:] - S_1[CY-1,:,:]) 
            
            
        a=35
            
        Total_Losses_tot=S_Env_Omega[a,:,:].sum()+ S_Env_Sigma[a,:,:].sum()+ S_Env_Phi[a,:].sum()+ S_Env_Theta[a,:].sum() + S_Gamma[a,:,:].sum()
        
        Total_Losses_lifestages = [S_Gamma[a,:,:].sum(),S_Env_Omega[a,:,:].sum(),S_Env_Phi[a,:].sum(),S_Env_Sigma[a,:,:].sum(),S_Env_Theta[a,:].sum()] 
        
        for r in range(0,Par_NoOfRegions):
            Total_Losses_regions[r]=round(((S_Env_Omega[a,r,:].sum()+ S_Env_Sigma[a,r,:].sum()+ S_Env_Phi[a,r].sum()+ S_Env_Theta[a,r].sum() + S_Gamma[a,r,:].sum())))

            Total_Losses_share[r]=Total_Losses_regions[r]/ Total_Losses_tot

            Total_Losses_Omega[r]   = S_Env_Omega[a,r,:].sum()
            Total_Losses_Sigma[r]   = S_Env_Sigma[a,r,:].sum()
            Total_Losses_Phi[r]     = S_Env_Phi[a,r].sum()
            Total_Losses_Theta[r]   = S_Env_Theta[a,r].sum()
            Total_Losses_Gamma[r]   = S_Gamma[a,r,:].sum()
            
            Losses=[S_Gamma[a,r,:].sum(),S_Env_Omega[a,r,:].sum(),S_Env_Phi[a,r].sum()+S_Env_Sigma[a,r,:].sum(),S_Env_Theta[a,r].sum()]
            
#            
            
                
            
        GN=[3]
        GN_Share_Losses_Omega = Total_Losses_Omega[GN].sum()
        GN_Share_Losses_Sigma = Total_Losses_Sigma[GN].sum()
        GN_Share_Losses_Phi = Total_Losses_Phi[GN].sum()  
        GN_Share_Losses_Theta = Total_Losses_Theta[GN].sum()
        GN_Share_Losses_Gamma = Total_Losses_Gamma[GN].sum()
        GN_in_use = S_1[35,GN,:].sum()
        
        
        GS=[0]
        GS_Share_Losses_Omega = Total_Losses_Omega[GS].sum()
        GS_Share_Losses_Sigma = Total_Losses_Sigma[GS].sum()
        GS_Share_Losses_Phi = Total_Losses_Phi[GS].sum()
        GS_Share_Losses_Theta = Total_Losses_Theta[GS].sum()
        GS_Share_Losses_Gamma = Total_Losses_Gamma[GS].sum()
        GS_in_use = S_1[35,GS,:].sum()
        
        GL=[7]
        GL_Share_Losses_Omega = Total_Losses_Omega[GL].sum()
        GL_Share_Losses_Sigma = Total_Losses_Sigma[GL].sum()
        GL_Share_Losses_Phi = Total_Losses_Phi[GL].sum()
        GL_Share_Losses_Theta = Total_Losses_Theta[GL].sum()
        GL_Share_Losses_Gamma = Total_Losses_Gamma[GL].sum()
        GL_in_use = S_1[35,GL,:].sum()
        
        
        
        
        results_share_GN=[Name_Scenario, 'Western Europe', GN_Share_Losses_Omega, GN_Share_Losses_Sigma ,GN_Share_Losses_Phi,GN_Share_Losses_Theta,GN_Share_Losses_Gamma,GN_in_use]
        results_share_GS=[Name_Scenario,'China', GS_Share_Losses_Omega, GS_Share_Losses_Sigma ,GS_Share_Losses_Phi,GS_Share_Losses_Theta,GS_Share_Losses_Gamma, GS_in_use]
        results_share_GL=[Name_Scenario,'Africa', GL_Share_Losses_Omega, GL_Share_Losses_Sigma ,GL_Share_Losses_Phi,GL_Share_Losses_Theta,GL_Share_Losses_Gamma, GL_in_use]

        results_regions = [Total_Losses_regions]
        results_in_use = [S_1[35,:,:].sum()]
            
            
        results_regions_GN_df = pd.DataFrame( results_share_GN)
        results_regions_GS_df = pd.DataFrame( results_share_GS)
        results_regions_GL_df = pd.DataFrame( results_share_GL)
        results_regions_df = pd.DataFrame(results_regions)
        results_in_use_df = pd.DataFrame(results_in_use)
        
       # writer=pd.ExcelWriter(Project_MainPath + 'General_Results\\' + 'Regional Results\\' + 'Pi_Chart_regions3.xlsx',engine='openpyxl')
           
                  
#        book=load_workbook(Project_MainPath + 'General_Results\\' + 'Regional Results\\' + 'Pi_Chart_regions3.xlsx')
#        
#        writer=pd.ExcelWriter(Project_MainPath + 'General_Results\\' + 'Regional Results\\' + 'Pi_Chart_regions3.xlsx',engine='openpyxl')
#        
#        writer.book=book
#        
#        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
#   
#        results_regions_GN_df.to_excel(writer, sheet_name='Pi_charts',startcol=int(Number_Scenario),startrow=1,index=False)
#        results_regions_GS_df.to_excel(writer, sheet_name='Pi_charts',startcol=int(Number_Scenario),startrow=10,index=False) 
#        results_regions_GL_df.to_excel(writer, sheet_name='Pi_charts',startcol=int(Number_Scenario),startrow=19,index=False) 
#        results_regions_df.to_excel(writer, sheet_name='Regional_losses',startcol=0,startrow=int(Number_Scenario)*2,index=False) 
#        results_in_use_df.to_excel(writer, sheet_name='Regional_losses',startcol=12,startrow=int(Number_Scenario)*2,index=False) 
#        writer.save()

Mylog.info('<p>Create plots.</p>')
        
#%% Plots       

if ScriptConfig['Time horizon'] == 2100 or 2300: 
    print('yes')
    if ScriptConfig['Modus'] == 'Trace 2015 mined copper':
                
                            
                
                fig = plt.figure()

 #               matplotlib.rc('font', size=10)


                plt.xlabel('Year')
                plt.ylabel('Copper mined in 2015 [kt]')
                x = list(range(0,86))
                Par_Time=list(range(2015,2101))
                colormap = plt.cm.gist_ncar
                plt.fill_between(Par_Time, 0, S_1[x,:,:].sum(axis=1).sum(axis=1) ,facecolor='#F5F5DC',label='Total copper stock')
                             

                #plot lines with sector stocks
                plt.plot(Par_Time, S_1[x,:,12:16].sum(axis=1).sum(axis=1),'b',label='Consumer & Electronics')
                plt.plot(Par_Time, S_1[x,:,0:5].sum(axis=1).sum(axis=1),'y',label='Building & Construction')
         #       plt.plot(Par_Time, S_1[x,:,0:5].sum(axis=1).sum(axis=1)-S_Brass[x,:,:].sum(axis=1).sum(axis=1),'--r')
                plt.plot(Par_Time, S_1[x,:,5:7].sum(axis=1).sum(axis=1),'g',label='Infrastructure')
                plt.plot(Par_Time, S_1[x,:,7:9].sum(axis=1).sum(axis=1),'r',label='Industrial')
                plt.plot(Par_Time, S_1[x,:,9:12].sum(axis=1).sum(axis=1),'c',label='Transport')
        
                #plot losses and fill between
                plt.fill_between(Par_Time, (S_1[x,:,:].sum(axis=1).sum(axis=1)), S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1),facecolor='whitesmoke',label='Obsolete stocks')
                plt.fill_between(Par_Time, S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1), S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1),facecolor='lightgrey',label='Dissipative losses')
                plt.fill_between(Par_Time, S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1), S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Phi[x,:].sum(axis=1),facecolor='darkgrey',label='Losses due to scrap and waste separation and sorting')
                plt.fill_between(Par_Time,  S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Phi[x,:].sum(axis=1), S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Phi[x,:].sum(axis=1)+S_Env_Theta[x,:].sum(axis=1),facecolor='dimgrey',label='Losses by copper recovery from scrap in recycling route')
                plt.fill_between(Par_Time,   S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Phi[x,:].sum(axis=1)+S_Env_Theta[x,:].sum(axis=1), S_1[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1)+S_Env_Phi[x,:].sum(axis=1)+S_Env_Theta[x,:].sum(axis=1)+ S_Gamma[x,:,:].sum(axis=1).sum(axis=1) ,facecolor='k',label='Not collected copper items')
                  
        
                #define shape of figure
                my_xticks = [Par_Time]
             #   plt.legend(frameon=True)
             #   plt.legend(bbox_to_anchor=(1.04,1), loc="lower center")
                xmin, xmax, ymin, ymax = 2015, 2100, 0, 20000
                plt.axis([xmin, xmax, ymin, ymax])
                plt.savefig(Path_Result +'stockplot.svg',acecolor='w',edgecolor='w', bbox_inches='tight')
                plt.savefig(Path_Result +'stockplot.png',acecolor='w',edgecolor='w', dpi=500)
        

#
#                Figure_2_paper =pd.DataFrame({#'Year':Par_Time, \
#                                              Name_Scenario + '_Consumer & Electronics':S_1[x,:,12:16].sum(axis=1).sum(axis=1), \
#                                              Name_Scenario + '_Building & Construction':S_1[x,:,0:5].sum(axis=1).sum(axis=1), \
#                                              Name_Scenario + '_Infrastructure':S_1[x,:,5:7].sum(axis=1).sum(axis=1), \
#                                              Name_Scenario + '_Industrial':S_1[x,:,7:9].sum(axis=1).sum(axis=1), \
#                                              Name_Scenario + '_Transport':S_1[x,:,9:12].sum(axis=1).sum(axis=1), \
#                                              Name_Scenario + '_Obsolete stocks':S_Env_Omega[x,:,:].sum(axis=1).sum(axis=1) ,\
#                                              Name_Scenario + '_Dissipative losses':S_Env_Sigma[x,:,:].sum(axis=1).sum(axis=1), \
#                                              Name_Scenario + '_Losses due to scrap and waste separation and sorting':S_Env_Phi[x,:].sum(axis=1), \
#                                              Name_Scenario + '_Losses by copper recovery from scrap in recycling route':S_Env_Theta[x,:].sum(axis=1), \
#                                              Name_Scenario + '_Not collected copper items':S_Gamma[x,:,:].sum(axis=1).sum(axis=1) 
#
#                                              })
#                                              
#                                              #,,,,,]
#                
#               # Figure_2_paper_df = pd.DataFrame(Figure_2_paper)
#                  
#
#                book=load_workbook('C:\\Users\\sklose\\Documents\\Writings\\Paper\\Paper_Number_1\\Submit_JIE\\supporting_information_1.xlsx')
#                
#                writer=pd.ExcelWriter('C:\\Users\\sklose\\Documents\\Writings\\Paper\\Paper_Number_1\\Submit_JIE\\supporting_information_1.xlsx',engine='openpyxl')
#                
#                writer.book=book
#                
#                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
#                
#                
#                Figure_2_paper.to_excel(writer, sheet_name='figure 2',startcol=34,startrow=0,index=False)
#                
#                
#                writer.save()       

                #calculate average lifetime of copper in the technosphere
                Figurecounter = 2
                plt.figure(2)
                
                x1=list(range(0,100))
                y= S_1[x,:,:].sum(axis=1).sum(axis=1)
                def func(x, a, b, c):
                    return a*np.exp(-c*(x-b))
                popt, pcov = curve_fit(func, x, y)
                #plot(x,y)
              #  plot(x1,func(x1,*popt))
               # show()
        
                #integrate underneath function!!
                area1 = trapz(y, dx=1)
                # Compute the area using the composite Simpson's rule.
                area2 = simps(y, dx=1)
#        
                x3=list(range(85,100000))
                area3 = simps(func(x3,*popt), dx=1)
                area=round(area1+area3)
                
               
                
          
#                plt.annotate('Average life-time of copper ='+ str(area),
#                    xy=(1.5, 0), xytext=(1, 5),
#                    xycoords=('axes fraction', 'figure fraction'),
#                    textcoords='offset points',
#                    size=14, ha='right', va='bottom')
                
                #calculate halftime (year when half of the copper is left in the anthroposphere)
                
                
                y100= S_1[0:86,:,:].sum()/19800
                Circ=round(y100/85,2)
                print(Circ)
                Total_copper=F_0_8[SY,:,:].sum()
                Half_copper=Total_copper/2
                
                Halftime=str(min(m for m in range (0,Par_NoOfYears) if S_1[m,:,:].sum(axis=0).sum(axis=0) < Half_copper))
                Halftime_year=int(Halftime) + 2015

                #calculate Amount of copper left afer 50 years
                Copper_50=round(S_1[50,:,:].sum(axis=0).sum(axis=0)/19800,2)
               
                #number of circles of a copper atom on average
                n = round(F_8_1[:,:,:].sum()/19800,2)
                
                n_regions =np.zeros((Par_NoOfRegions))
                area_regions = np.zeros((Par_NoOfRegions))
                for r in range(0,Par_NoOfRegions):
                    n_regions[r] = round(F_8_1[:,r,:].sum()/F_0_8[:,r,:].sum(),2)
                    
                    y_regions= S_1[x,r,:].sum(axis=1)/F_0_8[:,r,:].sum()
                    x1=list(range(0,100))
                    popt, pcov = curve_fit(func, x, y_regions)
      #              plot(x,y)
        #            plot(x1,func(x1,*popt))
          #          show()
         #       
                    #integrate underneath function!!
                    area1 = trapz(y_regions, dx=1)
               
                    x3=list(range(85,100000))
                    area3 = simps(func(x3,*popt), dx=1)
                    area=round(area1+area3)
                    area_regions[r]=area
                #calculate Amount of copper left afer 50 years
             # Copper_100=round(S_1[100,:,:].sum(axis=0).sum(axis=0))
               
                
                results_sensitivity=[Number_Scenario, Name_Scenario,area,n,Copper_50, Circ]

                results_sensitivity_df = pd.DataFrame(results_sensitivity)
              
                book=load_workbook(Project_MainPath + 'General_Results\\' +'Results_Sensitivity.xlsx')
                
                writer=pd.ExcelWriter(Project_MainPath + 'General_Results\\' +'Results_Sensitivity.xlsx',engine='openpyxl')
                
                writer.book=book
                
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                
                
                results_sensitivity_df.to_excel(writer, sheet_name='Sheet1',startcol=int(Number_Scenario),startrow=1,index=False)
                
                
                writer.save()

                EoLRR_min=min(EoLRR[:,:].sum(axis=1))      

                a=np.round(S_1[:,:,:].sum(axis=1).sum(axis=1),decimals=0)
                
                TotalStockInTime1=[Number_Scenario, Name_Scenario]
                
                TotalStockInTime1.extend(a)
                
                TotalStockInTime_df = pd.DataFrame(TotalStockInTime1)
                              
                book=load_workbook(Project_MainPath + 'General_Results\\' +'Results_totalstock.xlsx')
                
                writer=pd.ExcelWriter(Project_MainPath + 'General_Results\\' + 'Results_totalstock.xlsx' ,engine='openpyxl')
                
                writer.book=book
                
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                
                
                TotalStockInTime_df.to_excel(writer, sheet_name='Sheet1',startcol=int(Number_Scenario),startrow=1,index=False)


                writer.save()


   
    
#    
#                Figurecounter = 3
#                
#                
#               # BAU_stocks[:]= S_1[:,:,:].sum(axis=1).sum(axis=1)
#                plt.figure(3)
#                
#                x = list(range(0,86))
#                Par_Time=list(range(2015,2101))
#            
#         #       plt.plot(Par_Time, S_1[x,:,:].sum(axis=1).sum(axis=1)/BAU_stocks[x],'k',label='BAU')
#                plt.plot(Par_Time, S_1[x,:,:].sum(axis=1).sum(axis=1)/BAU_stocks[x],'b',label='')
#           #     plt.plot(Par_Time, S_1[x,:,:].sum(axis=1).sum(axis=1)/BAU_stocks[x],'k:',label='\theta_inf')
#         #       plt.plot(Par_Time, S_1[x,:,:].sum(axis=1).sum(axis=1)/BAU_stocks[x],'darkgrey',marker='o',markerfacecolor='None',markevery=3,linestyle='', label='$\gamma$$_inf$')
#       #         plt.plot(Par_Time, S_1[x,:,:].sum(axis=1).sum(axis=1)/BAU_stocks[x],'darkgrey',marker='+',markerfacecolor='None',markevery=3,linestyle='', label='$\phi$$_{inf}$')
#
#                my_xticks = [Par_Time]
#                
#                                
#                plt.legend(frameon=True)
#                plt.legend(bbox_to_anchor=(1.04,1), loc="lower center")
#                xmin, xmax, ymin, ymax = 2015, 2100, 0.6, 1.42
#                print(xmin, xmax, ymin, ymax)
#                plt.axis([xmin, xmax, ymin, ymax])
#                plt.savefig(Path_Result +'sensStockplot.svg',acecolor='w',edgecolor='w', bbox_inches='tight')
#                plt.savefig(Path_Result +'sensStockplot.png',acecolor='w',edgecolor='w', bbox_inches='tight', dpi=500)
#                
#workbook  = Sensitivity_Workbook.book
#worksheet = Sensitivity_Workbook.sheets['Sheet1']
#
#chart=workbook.add_chart({'type':'line'})
## Configure the series of the chart from the dataframe data.
#chart.add_series({'values': 'results_sensitivity_df'})
#
## Insert the chart into the worksheet.
#worksheet.insert_chart('D2', chart)
## Create a chart object.
#chart = workbook.add_chart({'type': 'column'})
#
#worksheet.insert_chart('D2', chart)

#workbook  = Sensitivity_Workbook.book
#worksheet = Sensitivity_Workbook.sheets['Sheet1']
#
#chart=workbook.add_chart({'type':'line'})
## Configure the series of the chart from the dataframe data.
#chart.add_series({'values': 'results_sensitivity_df'})
#
## Insert the chart into the worksheet.
#worksheet.insert_chart('D2', chart)
## Create a chart object.
#chart = workbook.add_chart({'type': 'column'})
#
#worksheet.insert_chart('D2', chart)



#Senstitivit_Pandas_Sheet=Sensitivity_Workbook.parse('Sheet1')

#writer = pd.ExcelWriter(Path_Data +'Results_Sensitivity.xlsx', engine='xlsxwriter')
#Senstitivit_Pandas_Sheet.to_excel(writer, 'Sheet1')
#writer.save()
#
#SW=load_workbook(Path_Data +'Results_Sensitivity.xlsx')
#print(SW.get_sheet_names())
# 
# 
#Sensitivity_Workbook = xlsxwriter.Workbook(Path_Data +'Results_Sensitivity.xlsx')
#Sensitivitysheet  = Sensitivity_Workbook.add_worksheet('Sheet1')
#
#
#.get_worksheet_by_name('Sheet1')
#
#Sensitivitysheet. 
#
#Sensitivity_Workbook.save('Sensitivitysheet')
#
#sheet_by_name('Sensitivity')
#Sensitivitysheet2 = Sensitivity_Workbook.add_worksheet('flow_data')
#Sensitivitysheet.write('B2' ,    "Source")
#Sensitivity_Workbook.

#ormat to use to highlight cells.
#               bold = workbook.add_format({'bold': True})
#    
#                #Add here the flows you want to have in the Sankey Diagram 
#                FlowData = (
#                ['Source','Value','Color', 'Flow_Style','Target'],
#                ['PP', round(Global_Copper_Production2015),'(200,191,55)','ab','M'],
#                ['M',round(F_8_1[SY+1:,:,0:5].sum()),'(0,191,255)','ab','BC'])
#
#
#                # Start from the first cell. Rows and columns are zero indexed.
#                row = 0
#                col = 0
#
#    # Iterate over the data and write it out row by row.
#    for Source,Value, Color,Flow_Style,Target in (FlowData):
#        Flowworksheet.write(row, col,     Source)
#        Flowworksheet.write(row, col + 1, Value)
#        Flowworksheet.write(row, col + 2, Color)


    
    
    
                             
    plt.show()

print(S_Gamma[35,:,:].sum())  
print(S_Env_Omega[35,:,:].sum())      
print(S_Env_Phi[35,:].sum())     
print(S_Env_Sigma[35,:,:].sum())           
print(S_Env_Theta[35,:].sum())         
           

    
print('END')    
    #