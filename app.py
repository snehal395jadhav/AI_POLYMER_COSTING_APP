"""
Molded Plastic Products - Costing Web Application
A secured web application replicating the Excel costing template.
"""
import os
import re
import json
import sqlite3
import hashlib
import secrets
from datetime import datetime, timedelta
from functools import wraps
from io import BytesIO
from urllib import error as urlerror
from urllib import request as urlrequest

