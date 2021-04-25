from distutils.core import setup
import py2exe
import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
import os
import os.path
from os import path
import glob
from configparser import ConfigParser
import pyodbc
import logging
import shutil

setup(name="downLVC_SII",
           version="1.0",
           description="Descarga de LV/LC desde sitio sii",
           author="RaYaKho",
           author_email="ryanezk@gmail.com",
           license="Commercial",
           scripts=['downlvc_sii.py'],
           console=['downlvc_sii.py'])
