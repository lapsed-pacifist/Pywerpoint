from setuptools import setup, find_packages

setup(name='pywerpoint',
        version='1.0',
        description='Easy to use tools for transfering data to powerpoint',
        author='Shaun Read',
        author_email='shaun.read@linkdex.com',
        packages=find_packages(exclude=['ez_setup', 'tests', 'tests.*']),
        package_data={'': ['slide_layout_enumeration.csv', 'table_temp.pptx']},
        include_package_data=True,
        install_requires=[],
)