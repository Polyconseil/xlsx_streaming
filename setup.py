from setuptools import setup, find_packages

PACKAGE = 'xlsx_streaming'
description = 'Export your data as an xlsx stream'


def get_long_description():
    with open('README.rst') as readme_file:
        return readme_file.read()


REQUIREMENTS = [
    'zipstream>=1.1.3',
]

setup(
    name=PACKAGE,
    version='1.0.0',
    description=description,
    long_description=get_long_description(),
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ],
    keywords=['xlsx', 'excel', 'streaming'],
    author='Polyconseil',
    author_email='opensource+%s@polyconseil.fr' % PACKAGE,
    url='https://github.com/Polyconseil/%s/' % PACKAGE,
    license='GNU GPLv3',
    packages=find_packages(exclude=['docs', 'tests']),
    setup_requires=[
        'setuptools',
    ],
    install_requires=REQUIREMENTS,
    test_suite='tests',
)
