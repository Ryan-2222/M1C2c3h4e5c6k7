@echo off
:start
cls

python ./get-pip.py
set python_ver=36

cd \
cd \python%python_ver%\Scripts\
pip install openpyxl
pip install plotly
pip install kaleido
pip install termcolor
pip install colorama

pause
exit
