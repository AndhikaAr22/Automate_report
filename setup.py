from setuptools import setup, find_packages
setup(name='report',
      version='0.1.0',
      packages=['report','report.src', 'report.src.operators'],
      install_requires=[
        'numpy',
        'pandas',
        'openpyxl',
        'discord==1.7.3'
    ]
    )