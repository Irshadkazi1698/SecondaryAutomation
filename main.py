from DSCValidationAutomation.MatchingFileCreation import *
from DSCValidationAutomation.BannerQCAutomation import *
import GUI

class BannerValidation():

    def __init__(self,InputDir,OutPutDir,BannerFileInfo,CountFileInfo,LongVarFileName,TabPlanFileName,TabPlanNumb,QuestionIndexInfo,LabelIndexInfo,BaseTextIndexInfo,SheetNameInfo):
        # These are the names we'll be using for our output Excel files.
        self.COUNT_FILE_NAME = CountFileInfo
        self.BANNER_FILE_NAME = BannerFileInfo
        self.UPDATED_MATCHING_FILE_NAME = 'Matched_Variables'
        self.SUMMARY_FILE_NAME = 'Unmatched_Summary'
        self.FINAL_COMPARISON_FILE_NAME = 'Final Comparison'
        self.LongVar = LongVarFileName
        self.TabPlan = TabPlanFileName
        self.TabPlanNumb = TabPlanNumb
        self.QuestionIndex = QuestionIndexInfo
        self.LabelIndex = LabelIndexInfo
        self.BaseTextInfo = BaseTextIndexInfo
        self.SheetName = SheetNameInfo

        # Here's where the script expects to find your input files and where it will save the results.
        self.INPUT_FOLDER_PATH = InputDir
        self.OUTPUT_FOLDER_PATH = OutPutDir

        # This is a specific marker the script looks for in the 'Counts' file to identify variable names.
        self.VARIABLE_PREFIX = 'VARIABLE_NAME = '


    def CreatingMatchingFileInOutput(self):

        return main_process(self.LongVar,
                            self.TabPlan,
                            self.INPUT_FOLDER_PATH,
                            self.OUTPUT_FOLDER_PATH,
                            self.BANNER_FILE_NAME,
                            self.TabPlanNumb,
                            self.QuestionIndex,
                            self.LabelIndex,
                            self.BaseTextInfo,
                            self.SheetName
                            )
    
    def BannerValidationAutomation(self):
       return  main(self.INPUT_FOLDER_PATH,
                    self.OUTPUT_FOLDER_PATH,
                    self.COUNT_FILE_NAME,
                    self.BANNER_FILE_NAME,
                    self.UPDATED_MATCHING_FILE_NAME,
                    self.FINAL_COMPARISON_FILE_NAME,
                    self.SUMMARY_FILE_NAME,
                    self.VARIABLE_PREFIX
                    )
    

# AddInputDirectory = input('''
#     Please add the Input directory. The Input Directory should contains below files \n

#         1. TabPlan \n
#         2. Banner (It should be rename to Banner) \n
#         3. Counts (It should be rename to Counts) \n
#         4. Numneric Variable List \n
#         5. Categorical VariablName file (It should be rename to Var_name) \n
# ''')


# AddOutputDirectory = input("Please add the Output Directory : ")

# AddNumericVariableFileName = input("Please Specify numeric Variable File Name without format extension : ")

# AddTabPlanFileName = input("Please Specify TabPlan File Name without format extension : ")

# print(f"Input Directory Info : {AddInputDirectory}\n")

# print(f'Output Directory Info : {AddOutputDirectory}\n')

# print(f"Numeric File Info : {AddNumericVariableFileName}\n")

# print(f"Tabplan file Info : {AddTabPlanFileName}\n")


# ValidationObj = BannerValidation(AddInputDirectory,
#                                  AddOutputDirectory,
#                                  AddNumericVariableFileName,
#                                  AddTabPlanFileName,
#                                  )

# #Call functions
# ValidationObj.CreatingMatchingFileInOutput()
# ValidationObj.BannerValidationAutomation()






