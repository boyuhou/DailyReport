import os
from setuptools import setup, find_packages

version_file = os.path.join(os.path.dirname(__file__), 'VERSION')
version = 'VERSION-NOT-FOUND'

if os.path.exists(version_file):
    with open(version_file) as version_file:
        version = version_file.read().strip()

setup(name='DailyReport',
      url='https://github.com/boyuhou/DailyReport',
      version=version,
      description='Python Version of Daily Report',
      maintainer='Bryan Hou',
      maintainer_email='boyuhou@gmail.com',
      packages=find_packages(),
      install_requires=[
          'pandas==0.24.2',
          'xlrd==1.2.0',
          'tables==3.4.4',
          'docxcompose==1.0.0a16',
          'comtypes==1.1.7',
          'matplotlib==2.2.2',
          'docxtpl==0.5.17',
          'python-docx==0.8.7',
          'colour==0.1.5',
      ],
      extras_require={
          'dev': [
            'jupyterlab',
            'click==7.0'
          ]
      },
      package_data={
         "DailyReport": [
            "template/DAILY_TEMPLATE.docx",
         ],
      },
)
