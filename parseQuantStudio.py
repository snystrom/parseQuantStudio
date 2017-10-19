#! /usr/bin/python3

import sys
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import math

def parseQuantStudioSheet(excelFile, sheetName = 0):
    data = pd.read_excel(excelFile, skiprows = 40, sheetname = sheetName)
    # replace spaces for _ in colnames
    data.columns = [c.replace(' ', '_') for c in data.columns]
    return(data)

def coerceDataTypes(data, strCols, intCols, floatCols):
    """
    data = pandas data frame
    <type>Cols = list of colnames to be converted to str/int/float respectively
    """
    #TODO: Allow empty values of these columns

    
    for strCol, intCol, floatCol in zip(strCols, intCols, floatCols):
        data[strCol] = [str(i) for i in data[strCol]]
        data[intCol] = [int(i) for i in data[intCol]]
        data[floatCol] = [float(i) for i in data[floatCol]]

    return(data)

def getSampleData(excelFile):
    """
    Extracts sample-specific information, to be joined with other data tables
    to make sense of QuantStudio's purposeful obfuscation.
    """
    sampleData = parseQuantStudioSheet(excelFile, "Sample Setup")
    sampleData = sampleData.loc[:, ['Well', 'Well_Position', 'Sample_Name', 'Sample_Color', 'Target_Name']]
    
    # Now remove anything with 3 NaN values == not defined sample
    sampleData.dropna(thresh=3, inplace=True)
    
    # convert all values to strings
    for col in sampleData.columns:
        sampleData[col] = [str(i) for i in sampleData[col]]

    return(sampleData)

def getAmplificationData(excelFile, sampleData):
    """
    Grabs amplification curve data & merges w/ sample-specific information
    """
    ampData = parseQuantStudioSheet(excelFile, "Amplification Data")
    ampData.dropna(thresh = 3, inplace = True)

    strCols = ['Well', 'Target_Name']
    intCols = ['Cycle']
    floatCols = ['Rn', 'Delta_Rn']

    ampData = coerceDataTypes(ampData, strCols, intCols, floatCols)
    
    # Annotate w/ sample-specific information
    ampDataAnno = ampData.merge(sampleData, on = 'Well', how = "inner")
    renameCols = {'Target_Name_x': 'Target_Name'} 
    ampDataAnno.rename(columns = renameCols, inplace = True)

    return(ampDataAnno)

def getMeltCurveData(excelFile, sampleData):
    """
    Grabs melt curve data & merges w/ sample-specific information
    """
    meltData = parseQuantStudioSheet(excelFile, "Melt Curve Raw Data")
    meltData.dropna(thresh = 3, inplace = True)

    strCols = ['Well', 'Well_Position']
    intCols = ['Reading']
    floatCols = ['Temperature', 'Fluorescence', 'Derivative']

    meltData = coerceDataTypes(meltData, strCols, intCols, floatCols)
    
    meltDataAnno = meltData.merge(sampleData, on = 'Well', how = "inner")
    renameCols = {'Well_Position_x': 'Well_Position'}
    meltDataAnno.rename(columns = renameCols, inplace = True)

    return(meltDataAnno)

def getResults(excelFile, sampleData):
    """
    Grabs sample-level results & merges w/ sample-specific information
    """
    resData = parseQuantStudioSheet(excelFile, "Results")
    resData.drop(resData.tail(4).index, inplace = True)

    strCols = ['Well', 'Well_Position', 'Sample_Name']
    intCols = ['Baseline_Start', 'Baseline_End']
    floatCols = ['RQ', 'RQ_Min', 'RQ_Max', 
            'CT', 'CT_Mean', 'CT_SD', 
            'Delta_Ct', 'Ct_Threshold',
            'Tm1', 'Tm2', 'Tm3']

    # Toss out all the other columns which contain mostly variables for working within QuantStudio
    resCols = strCols + intCols + floatCols
    resData = resData.loc[:, resCols]

    resData.dropna(how = "all", inplace=True)
    
    resData = coerceDataTypes(resData, strCols, intCols, floatCols)
    
    resDataAnno = resData.merge(sampleData, on = 'Well', how = "inner")
    renameCols = {'Well_Position_x': 'Well_Position', 'Sample_Name_x' : 'Sample_Name'}
    resDataAnno.rename(columns = renameCols, inplace = True)

    return(resDataAnno)

def plotCT(resultsData):
    """
    returns barplot of CT values for each sample & target
    """
    #TODO: Error checking for valid cols
    sns.set_context("poster")
    CT = sns.factorplot(data = resultsData,
        x = "Target_Name", y = "CT", hue = "Sample_Name", 
        kind = "bar", size = 8, aspect = 2, legend = False)
    CT.ax.legend(loc = 2, bbox_to_anchor=(1.05, 1), borderaxespad=0.)
    return(CT)

def plotAmplificationCurve(ampData):
    """
    Single plot of all amplification curve data
    """
    #TODO: Error checking for valid cols
    sns.set_context("poster", font_scale = 2)
    ampCurve = sns.FacetGrid(ampData, 
        col = "Target_Name", row = "Sample_Name", hue = "Sample_Name",
        size = 7.5, aspect = 2)
    ampCurve = ampCurve.map(plt.scatter, "Cycle", "Rn")
    return(ampCurve)

def plotMeltCurve(meltCurveData):
    """
    Single plot of all melt curve data
    """
    #TODO: Error checking for valid cols
    sns.set_context("poster", font_scale = 2)
    meltCurve = sns.FacetGrid(meltCurveData, row = "Target_Name", hue = "Sample_Name",
        size = 7.5, aspect = 2)
    meltCurve = meltCurve.map(plt.scatter, "Reading", "Derivative")
    return(meltCurve)
    
def main():
    """
    Usage: python3 parseQuantStudio.py <QuantStudioOutput>.xls
    Requires the following tables to be output:
    """
    excelFile = sys.argv[1]
    sampleData = getSampleData(excelFile)
    
    # Amplification curve data:
    ampData = getAmplificationData(excelFile, sampleData)
    
    # Melt curve data: 
    meltCurveData = getMeltCurveData(excelFile, sampleData)
    
    # Results
    resData = getResults(excelFile, sampleData)
    
    # Make CT plot
    CT = plotCT(resData) 
    CT.savefig('CT_Values.png')

    # Amplification Plot
    ampCurve = plotAmplificationCurve(ampData)
    ampCurve.savefig("Amplification_Curves.png")
    
    # MeltCurve Plot
    meltCurve = plotMeltCurve(meltCurveData)
    meltCurve.savefig("Melt_Curves.png")

    pass

if __name__ == "__main__":
    main()



