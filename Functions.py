from recurdyn import *
# from recurdyn import Chart
# from recurdyn import MTT2D
# from recurdyn import FFlex
# from recurdyn import RFlex
from recurdyn import Tire
import numpy as np
import glob
import os
import subprocess
import time
import Var
import re
import pandas as pd
import shutil
import matplotlib.pyplot as plt


####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
# DO NOT MODIFY BELOW CODES
# DO NOT MODIFY BELOW CODES
# DO NOT MODIFY BELOW CODES
# DO NOT MODIFY BELOW CODES
# DO NOT MODIFY BELOW CODES
# DO NOT MODIFY BELOW CODES
####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
# Common Variables
rdSolverDir = "\"C:\Program Files\FunctionBay, Inc\RecurDyn V9R5\Bin\Solver\RDSolverRun.exe\""
app = None
application = None
model_document = None
plot_document = None
model = None

ref_frame_1 = None
ref_frame_2 = None

# initialize() should be called before ProcessNet function call.
def initialize():
    global app
    global application
    global model_document
    global plot_document
    global model
    
    app = dispatch_recurdyn()
    application = IApplication(app.RecurDynApplication)
    model_document = application.ActiveModelDocument
    if model_document is not None:
        model_document = IModelDocument(model_document)
    plot_document = application.ActivePlotDocument
    if plot_document is not None:
        plot_document = IPlotDocument(plot_document)
    
    if model_document is None and plot_document is None:
        application.PrintMessage("No model file")
        model_document = application.NewModelDocument("Examples")
    if model_document is not None:
        model_document = IModelDocument(model_document)
        model = ISubSystem(model_document.Model)
    
    return application, model_document, plot_document, model

# dispose() should be called after ProcessNet function call.
def dispose():
    global application
    global model_document
    
    model_document = application.ActiveModelDocument
    if model_document is not None:
        model_document = IModelDocument(model_document)
    else:
        return
    
    if not model_document.Validate():
        return
    # Redraw() and UpdateDatabaseWindow() can take more time in a heavy model.
    # model_document.Redraw()
    # model_document.PostProcess() # UpdateDatabaseWindow(), SetModified()
    # model_document.UpdateDatabaseWindow()
    # If you call SetModified(), Animation will be reset.
    # model_document.SetModified()
    model_document.SetUndoHistory("Python ProcessNet")
####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
# DO NOT MODIFY ABOVE CODES
# DO NOT MODIFY ABOVE CODES
# DO NOT MODIFY ABOVE CODES
# DO NOT MODIFY ABOVE CODES
# DO NOT MODIFY ABOVE CODES
# DO NOT MODIFY ABOVE CODES
####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
####################################################################################################################################################
rdSolverDir = "\"C:\Program Files\FunctionBay, Inc\RecurDyn V9R5\Bin\Solver\RDSolverRun.exe\""

def CreateDir(DirectoryPath: str):
    """
    Creates a new directory.
    :param DirectoryPath: New directory path.
    :return:
    """
    if not os.path.exists(DirectoryPath):
        os.makedirs(DirectoryPath)
        print(f"Created new directory: {DirectoryPath}")
    else:
        pass

def Sec2Time(seconds: float):
    """
    Converts time seconds to Hours/Minutes/Seconds
    :param seconds:
    :return: Hours/Minutes/Seconds
    """
    Hr = int(seconds // 3600)
    seconds -= Hr * 3600
    Min = int(seconds // 60)
    seconds -= Min * 60
    Sec = seconds
    return Hr, Min, int(Sec)

def ChangePVvalue(model, PVname: str, PVvalue: float):
    """
    Change ParametricValue(PV) value in the model.
    :param model: Target model.
    :param PVname: Name of target PV.
    :param PVvalue: New value for the PVname.
    :return:
    """
    PV = IParametricValue(model.GetEntity(PVname))
    PV.Value = PVvalue

def Import(importfile: str):
    """
    Import file in the RecurDyn.
    :param importfile: Target file, can be CAD/Data/FlexibleBodies/rdyn...anything RecurDyn supports.
    :return:
    """
    modelPath = model_document.GetPath(PathType.WorkingFolder)
    print(f"Imported file: {importfile}")
    model_document.FileImport(importfile)

def ExportSolverFiles(OutputFolderName: str,
        OutputFileName: str,
        EndTime: int = 1,
        NumSteps: int = 101,
        PlotMultiplierStepFactor: int = 1):
    """
    Exports *.rmd, *.rss, and copies *.(DependentExt) to directory modelPath/OutputFolderName/*.* for batch automated solving.
    DependentExt may include tire files (*.tir), GRoad files (*.rdf), or flexible meshes.
    :param OutputFolderName:
    :param OutputFileName:
    :param EndTime: Simulation end time
    :param NumSteps: Simulation steps
    :param PlotMultiplierStepFactor:
    :return:
    """
    model_document = application.ActiveModelDocument
    model = model_document.Model
    modelPath = model_document.GetPath(PathType.WorkingFolder)
    print(f"{modelPath}{OutputFolderName}\\{OutputFileName}")
    CreateDir(f"{modelPath}{OutputFolderName}\\{OutputFileName}")
    # Copy dependency files
    DependentExt = ("tir", "rdf",)
    DependentFiles = []
    for ext in DependentExt:
        DependentFiles.extend(glob.glob(f"{modelPath}*.{ext}"))
    for file in DependentFiles:
        shutil.copy(file, f"{modelPath}{OutputFolderName}\\{OutputFileName}")
    # Analysis Property
    model_document.ModelProperty.DynamicAnalysisProperty.MatchSolvingStepSize = True
    model_document.ModelProperty.DynamicAnalysisProperty.MatchSimulationEndTime = True
    model_document.ModelProperty.DynamicAnalysisProperty.PlotMultiplierStepFactor.Value = PlotMultiplierStepFactor
    # RMD export
    model_document.FileExport(f"{modelPath}{OutputFolderName}\\{OutputFileName}\\{OutputFileName}.rmd",
                              True)
    # RSS export
    RSScontents = f"SIM/DYN, END = {EndTime}, STEP = {NumSteps}\nSTOP"
    rss = open(f"{modelPath}{OutputFolderName}\\{OutputFileName}\\{OutputFileName}.rss", 'w')
    rss.write(RSScontents)
    rss.close()

def WriteBatch(SolverFilesFolderName: str, parallelBatches: 1):
    """
    Write *.bat execution files for batch solving.
    :param SolverFilesFolderName:
    :param parallelBatches:
    :return:
    """
    global rdSolverDir
    application.ClearMessage()
    model_document = application.ActiveModelDocument
    model = model_document.Model
    modelPath = model_document.GetPath(PathType.WorkingFolder)
    RMDlist = glob.glob(f"{modelPath}{SolverFilesFolderName}\\**\\*.rmd", recursive=True)
    for i in range(parallelBatches):
        BatchFileName = f"{SolverFilesFolderName}_{i + 1}.bat"
        interval = round(len(RMDlist) / parallelBatches)
        bat = open(f"{modelPath}{SolverFilesFolderName}\\{BatchFileName}", 'w')
        if i == parallelBatches - 1:  # Last index
            idx_start = i * interval
            idx_end = len(RMDlist)
        else:
            idx_start = i * interval
            idx_end = (i + 1) * interval
        for rmdName in RMDlist[idx_start: idx_end]:
            solverfilename = os.path.basename(rmdName).split('.')[:-1]
            if len(os.path.basename(rmdName).split('.')) > 2:
                solverfilename = '.'.join(os.path.basename(rmdName).split('.')[:-1])
            else:
                solverfilename = ''.join(solverfilename)
            BATcontent = []
            BATcontent.append(modelPath[:2])  # Drive Name
            BATcontent.append(f"cd {os.path.dirname(rmdName)}")  # cd RMD path
            BATcontent.append(f"{rdSolverDir} {solverfilename} {solverfilename}")  #
            bat.writelines(line + "\n" for line in BATcontent)
        bat.close()
        application.PrintMessage(f"Created batch executable {modelPath}{SolverFilesFolderName}\\{BatchFileName}")

def RPLT2CSV(SolverFilesPath: str):
    """
    Scans and reads all *.rplt files in the SolverFilesAbsPath (recursively scanned).
    Then export values from Var.DataExportTargets in *.csv format.
    :param SolverFilesPath: rplt files directory
    :return:
    """
    application.CloseAllPlotDocument()
    CSVExportDir = os.path.abspath(SolverFilesPath)
    CreateDir(CSVExportDir)
    RPLTlist = glob.glob(f"{CSVExportDir}\\**\\*.rplt", recursive=True)
    StartTime = time.time()
    for idx_rplt, rplt in enumerate(RPLTlist):
        application.NewPlotDocument("PlotDoc")
        application.OpenPlotDocument(rplt)
        plot_document = application.ActivePlotDocument
        rpltname = os.path.basename(rplt).split('.')[:-1]
        if len(os.path.basename(rplt).split('.')) > 2:
            rpltname = '.'.join(os.path.basename(rplt).split('.')[:-1])
        else:
            rpltname = ''.join(rpltname)
        DataExportTargets = [f"{rpltname}/{target}" for target in Var.DataExportTargets]
        CSVpath = f"{CSVExportDir}\\{rpltname}.csv"
        plot_document.ExportData(CSVpath, DataExportTargets, True, False, 8)
        application.ClosePlotDocument(plot_document)
        application.PrintMessage(f"Data exported {CSVpath} ({idx_rplt + 1}/{len(RPLTlist)})")
        print(f"Data exported({idx_rplt + 1}/{len(RPLTlist)}) {CSVpath} ")
    EndTime = time.time()
    Elapsed = EndTime - StartTime
    Hr, Min, Sec = Sec2Time(Elapsed)
    print(f"Finished, data export time: {Hr}hr {Min}min {Sec:.2f}sec")

def GenerateBatchSolvingDOE(TopFolderName: str, NumParallelBatches: int = 1, NumCPUCores: int = 16):
    """
    Creates batch-solving DOEs.
    :param TopFolderName: A new directory for batch solving files.
    :param NumParallelBatches: Number of executable *.bat files.
            NOTE: One execution of *.bat uses one RecurDyn/Professional license.
    :param NumCPUCores: Number of CPU threads for solving. 'AUTO' if 0. [1,2,4,8,16] are supported values.
    :return:
    """
    application.ClearMessage()
    if NumCPUCores:
        application.Settings.AutoCoreNumber = False
        application.Settings.CoreNumber = NumCPUCores
    else:
        application.Settings.AutoCoreNumber = True
    
    model_document = application.ActiveModelDocument
    modelPath = model_document.GetPath(PathType.WorkingFolder)
    model = model_document.Model
    EndTime = 1
    NumSteps = 100
    
    Counter = 1
    SamplePV = np.logspace(-2, 10, 3, endpoint=True)
    for samplepv in SamplePV:
        ChangePVvalue(model, "PV_SampleK", samplepv)
        SubFolderName = f"{TopFolderName}_{Counter:04d}"
        ExportSolverFiles(TopFolderName, SubFolderName, EndTime=EndTime, NumSteps=NumSteps)
        Counter += 1
    WriteBatch(TopFolderName, NumParallelBatches)

def RunDOE(TopFolderName: str, NumCPUCores: int = 16, EndTime: float = 1, NumSteps: int = 100):
    """
    Run automated simulations in GUI interface solver.
    :param TopFolderName: A new directory for solved files.
    :param NumCPUCores: Number of CPU threads for solving. 'AUTO' if 0. [1,2,4,8,16] are supported values.
    :return:
    """
    ############################# INITIAL SETTING #############################
    ############################# INITIAL SETTING #############################
    ############################# INITIAL SETTING #############################
    application.ClearMessage()
    model_document = application.ActiveModelDocument
    modelPath = model_document.GetPath(PathType.WorkingFolder)
    model = model_document.Model
    application.Settings.CreateOutputFolder = False
    model_document.UseOutputFileName = True
    model_document.ModelProperty.DynamicAnalysisProperty.SimulationStep.Value = NumSteps
    model_document.ModelProperty.DynamicAnalysisProperty.SimulationTime.Value = EndTime
    model_document.ModelProperty.DynamicAnalysisProperty.MatchSolvingStepSize = True
    model_document.ModelProperty.DynamicAnalysisProperty.MatchSimulationEndTime = True
    if NumCPUCores:
        application.Settings.AutoCoreNumber = False
        application.Settings.CoreNumber = NumCPUCores
    else:
        application.Settings.AutoCoreNumber = True
    ############################# INITIAL SETTING #############################
    ############################# INITIAL SETTING #############################
    ############################# INITIAL SETTING #############################
    Counter = 1
    SamplePV = np.logspace(-2, 10, 3, endpoint=True)
    for samplepv in SamplePV:
        ChangePVvalue(model, "PV_SampleK", samplepv)
        SubFolderName=f"{TopFolderName}_{Counter:04d}"
        model_document.OutputFileName = f"{TopFolderName}\\{SubFolderName}\\{SubFolderName}"
        model_document.Analysis(AnalysisMode.Dynamic)
        Counter += 1

if __name__ == '__main__':
    application, model_document, plot_document, model = initialize()
    #
    # Open SampleModel.rdyn and run code
    
    # For GUI solver,
    RunDOE("SampleDOE_GUI", 16) # This is equivalent to: GenerateBatchSolvingDOE("SampleDOE_GUI", 1, 16)
    RPLT2CSV("SampleDOE_GUI")  # Export CSV
    
    # For batch solver,
    GenerateBatchSolvingDOE("SampleDOE_Batch", 1, 16)
    # Run *.bat files and then run RPLT2CSV("SampleDOE_Batch")
    RPLT2CSV("SampleDOE_Batch")  # Export CSV
    #
    
    dispose()
