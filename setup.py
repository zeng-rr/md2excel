from setuptools import setup

setup(
    name='md2excel',
    packages=['md2excel'],
    package_dir={'md2excel': 'src/md2excel'},
    include_package_data=True,
    version='0.1.0',
    author='zeng-rr',
    author_email='nern-zeng@outlook.com',
    description='Markdown表格转Excel转换工具',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/zeng-rr/md2excel',
    install_requires=[
        'openpyxl>=3.0.0',
    ],
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
    entry_points={
        'console_scripts': [
            'md2excel=md2excel.cli:main',
        ],
    },
)