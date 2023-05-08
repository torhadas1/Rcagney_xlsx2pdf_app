import streamlit as st
import pandas as pd
import openpyxl
import win32com.client
import pythoncom
import io
import base64
import sys
import streamlit.web.cli as stcli


if __name__ == "__main__":
    sys.argv=["streamlit", "run", "final_app.py", "--global.developmentMode=false"]
    sys.exit(stcli.main())