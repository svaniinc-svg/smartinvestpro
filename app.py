import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt
from datetime import datetime
import re
from pathlib import Path
import uuid

st.set_page_config(
    page_title="Real Estate ROI Calculator",
    layout="wide"
)

st.title("Know the Deal Before You Invest")

st.caption(
    "No signup required. Run the numbers first. "
    "Join the investor deals list later if you want Charlotte-area opportunities "
    "and deal review follow-up."
)
