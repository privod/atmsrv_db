from setuptools import setup, find_packages
from os.path import join, dirname

import atmsrv_db

setup(
    name='atmsrv_db',
    version=atmsrv_db.__version__,
    packages=find_packages(),
    long_description=open(join(dirname(__file__), 'README.rst')).read(),

    install_requires=[
        'cx_Oracle',
        'xlwt',
        'rels',
        # 'watchdog',
        # 'progressbar2',
        # 'PyQt5',
    ],

    entry_points={
        'console_scripts': [
            'actual_ncr = atmsrv_db.order_reports:actual_ncr',
        ]
    },
    # include_package_data=True,
    test_suite='tests',
)