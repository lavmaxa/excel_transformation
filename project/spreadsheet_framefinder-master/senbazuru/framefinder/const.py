'''
Created on Sep 26, 2012

@author: cz
'''
import os
from pathlib import Path

#p = os.path.abspath( os.path.dirname( __file__ ) )
p = Path(__file__).parent.parent

# directory to store the original spreadsheets
_sheetdir = os.path.join(p,'data', 'testsheets')

# directory to store the output:
# each spreadsheet labeled with semantic labels for each row
_outputdir = os.path.join(p,'data', 'predictsheets')

# directories to store intermediate results
_crftempdir = os.path.join(p,'data', 'temp')
_crffeadir = os.path.join(_crftempdir, 'crf_fea')
_crfpredictdir = os.path.join(_crftempdir, 'crf_predict')

# template file for CRF++ to parse the provided features
_crfpptemplatepath = os.path.join(p,'data','template')
# training data provided for 100 spreadsheets downloaded from http://www.census.gov/
_crftraindatapath = os.path.join(p,'data','saus_train.data')

##################################################
# please specifiy the directory of CRF++
##################################################
# directory of installed CRF++
_crfppdir = os.path.abspath('C:\project\CRF++-0.54')


