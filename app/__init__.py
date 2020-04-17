import os
from flask import Flask

def create_app():
   app=Flask(__name__)
   app.config.from_mapping(
      SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev_key'
   )
   return app


# Initialize the app
app = Flask(__name__, instance_relative_config=True)

# Load the views
from app import views

# Load the config file
app.config.from_object('config')
