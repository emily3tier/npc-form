import json
import base64
import io
import urllib.request
import urllib.parse
import os
import datetime
from openpyxl import load_workbook

CLIENT_ID   = 'd6dd88e1-a49e-4350-8339-f0f42c4b3b2e'
TENANT_ID   = '667afa82-1126-4a78-8f76-0918c7f2a845'
BASE_FOLDER = 'UPC Submissions Automated'
TEMPLATE_NAME = 'NPC_Form_2026_1.xlsx'
