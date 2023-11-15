
import constants as c
import functions as f

#pip3 install pandas numpy matplotlib seaborn sklearn-features vaderSentiment python-pptx openpyxl nltk

chart1, chart2, chart3 = f.textAnalysis() 

passToPresentation = [chart1, chart2, chart3]

f.createExport(passToPresentation)
