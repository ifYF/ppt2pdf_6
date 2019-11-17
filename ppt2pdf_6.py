import comtypes.client
import os

def init_powerpoint():
   powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
   powerpoint.Visible = 1
   return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32, ExportAsFixedFormat = 2):
   if outputFileName[-4:] != 'pptx':
      outputFileName = outputFileName[:-6]
   if outputFileName[-3:] != 'ppt':
      outputFileName = outputFileName[:-5]
   if outputFileName[-3:] != 'pdf':
      outputFileName = outputFileName + ".pdf"
   # deck = powerpoint.Presentations.Open(inputFileName)
   deck = powerpoint.Presentations.Open(inputFileName, -1)
   # deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
   deck.ExportAsFixedFormat(outputFileName, ExportAsFixedFormat, OutputType = 4)
   # OutputType Value: 4->6张ppt讲义 8->4张ppt讲义 9->9张ppt讲义
   deck.Close()

def convert_files_in_folder(powerpoint, folder):
   files = os.listdir(folder)
   pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
   for pptfile in pptfiles:
       fullpath = os.path.join(cwd, pptfile)
       ppt_to_pdf(powerpoint, fullpath, fullpath)

if __name__ == "__main__":
   powerpoint = init_powerpoint()
   cwd = os.getcwd()
   convert_files_in_folder(powerpoint, cwd)
   powerpoint.Quit()