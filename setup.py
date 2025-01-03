from setuptools import setup, find_packages

setup(
    name="excel_analyzer",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "openpyxl>=3.0.0",
    ],
    entry_points={
        'console_scripts': [
            'excel-analyzer=src.cli:main',
        ],
    },
) 