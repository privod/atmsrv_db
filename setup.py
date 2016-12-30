from setuptools import setup, find_packages
from os.path import join, dirname

import atmsrv_db

setup(
    name='atmsrv_db',
    version=atmsrv_db.__version__,
    packages=find_packages(),
    long_description=open(join(dirname(__file__), 'README.txt')).read(),

    # install_requires=[
    #     'watchdog',
    #     'progressbar2',
    #     'PyQt5',
    # ],

    # entry_points={
    #     'console_scripts': [
    #         'atm-start = atmsrv_db.cheque_replace:count_by_month',
    #     ]
    # },
    # include_package_data=True,
    test_suite='tests',
)